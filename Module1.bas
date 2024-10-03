Attribute VB_Name = "Module1"
Sub Prime(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Sheets(nameOfSheet1).Activate
    ActiveSheet.AutoFilterMode = False
    txtCol1 = "Город"
    txtCol2 = "Тип"
    txtCol3 = "Размер"

    Set XCell = Sheets(nameOfSheet1).Cells.Find(txtCol1)
    Set YCell = Sheets(nameOfSheet1).Cells.Find(txtCol2)
    Set ZCell = Sheets(nameOfSheet1).Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
    
    '------создаем ключ типа---------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
   
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 2).Value = "1.8x1.2" Or _
            Cells(i, YCol + 2).Value = "1.86x1.27" Or _
            Cells(i, YCol + 2).Value = "1.7x1.2" Or _
            Cells(i, YCol + 2).Value = "1.84x1.27" Or _
            Cells(i, YCol + 2).Value = "1.7x1.1" Or _
            Cells(i, YCol + 2).Value = "1.75x1.15" _
            And Cells(i, YCol + 1).Value = "скролл" Or Cells(i, YCol + 1).Value = "Сити-лайт") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призма") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else If (Cells(i, YCol + 1).Value = "Скролл" And Cells(i, YCol + 2).Value = "3x6") _
            Then Cells(i, YCol).Value = "всад" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------создаем стоиомость own------

    Columns(16).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 16) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f3").Copy
    Range(Cells(2, 16), Cells(lLastRow, 16)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '----------преобразование в числа--------
    With ActiveSheet.UsedRange
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With
    
    '------город---------
    
Const ColtoFilter1 As Integer = 4
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 7
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("j2:j10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 9
    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("D1:D40")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 15
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("B2:B4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
'Set StartCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Range(Cells(XRow, 1))
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
        '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 6).Value = Cells(i - 1, 6).Value And Cells(i, 10).Value = Cells(i - 1, 10).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Bigmedia(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)

Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer
    
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Sheets(nameOfSheet1).Activate
    ActiveSheet.AutoFilterMode = False
    txtCol1 = "Город"
    txtCol2 = "Сеть"

    Set XCell = Sheets(nameOfSheet1).Cells.Find(txtCol1)
    Set YCell = Sheets(nameOfSheet1).Cells.Find(txtCol2)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    
    '--------формат для фильтра----------
    'Cells.MergeCells = False 'убрать объединение ячеек
    'Range("A1:L1").Select
    'Selection.Copy
    'Range("A2").Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    '    :=False, Transpose:=False
    'Rows("1:1").Select
    'Application.CutCopyMode = False
    'Selection.Delete Shift:=xlUp
  
    '------создаем ключ для типа---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призмавижн" Then Cells(i, YCol).Value = "биллборд" Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '---------исправлем название города---------
    
    Columns(XCol).Select
    Selection.Replace What:="Днипро", Replacement:="Днепр", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    '------город---------
    
Const ColtoFilter1 As Integer = 3
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 8

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("K2:K10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 6

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("E1:E6")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 18

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("C2:C3")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a2")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        'конец переноса ширины столбцов
    End With
    
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row

        '-------------------удалить дубликаты--------------------

    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value = Cells(i - 1, 3).Value And Cells(i, 4).Value = Cells(i - 1, 4).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----------------добавляем себестоимость------------------
    txtCol = "Price"  ' метка для столбца
    Set ZCell = ActiveSheet.Cells.Find(txtCol)
    If ZCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Price Bigmedia, "
    Else
    ZCol = ZCell.Column
    ZRow = ZCell.Row
    Columns(ZCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, ZCol).Select
    Cells(1, ZCol) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, ZCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = "Скролл" And Cells(i, 3).Value = "Киев" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB6") * Cells(i, ZCol + 1) _
            Else If Cells(i, 8).Value = "Скролл" And Cells(i, 3).Value = "Харьков" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB8") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "Скролл" And Cells(i, 3).Value = "Одесса" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB10") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "Скролл" And (Cells(i, 3).Value <> "Киев" Or Cells(i, 2).Value <> "Одесса" Or Cells(i, 2).Value <> "Харьков") _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB4") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "Ситилайт" And Cells(i, 3).Value = "Киев" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB7") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "Ситилайт" And Cells(i, 3).Value = "Харьков" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB9") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "Ситилайт" And Cells(i, 3).Value = "Одесса" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB11") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "Ситилайт" And (Cells(i, 3).Value <> "Киев" Or Cells(i, 3).Value <> "Одесса" Or Cells(i, 2).Value <> "Харьков") _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB5") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "биллборд" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("Скидки").Range("AB3") * Cells(i, ZCol + 1)
    Next
    End If

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Octagon(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

'---------убираем старые данные-----------
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet2).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "Город"
    txtCol2 = "Формат"

    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol1)
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    XCol = XCell.Column
    YCol = YCell.Column
    
    '------создаем ключ типа---------
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "Ситилайт 1.2х1.8[CF]" Or Cells(i, YCol + 1).Value = "Скроллер 1.2х1.8 [CF]") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит 3x6 [BB]" Or Cells(i, YCol + 1).Value = "Призматрон 3х6 [BB]") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else If Cells(i, YCol + 1).Value = "Скроллер 2.3х3.14 [BO]" _
            Then Cells(i, YCol).Value = "скролл" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
  
    
    '-------добавляем столбец для себестоиомости------
    Columns(10).Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    Cells(1, 10) = "Себестоимость"
        
    '----------преобразование в числа--------
    With ActiveSheet.UsedRange.Columns(16)
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With
    
    '------город---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 4
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("l2:l10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 21
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("d2:d4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
        '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 7).Value = Cells(i - 1, 7).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----------------добавляем себестоимость------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, 10).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = "биллборд" _
            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("Скидки").Range("AM3") * Cells(i, 11) _
            Else: If Cells(i, 4).Value = "ситилайт" _
            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("Скидки").Range("AM4") * Cells(i, 11) _
            Else Cells(i, 10).Value = ThisWorkbook.Worksheets("Скидки").Range("AM5") * Cells(i, 11)
    Next



    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub SVO_news(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)

Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer
    
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    txtCol1 = "Регион"
    txtCol2 = "Конструкция"
    txtCol3 = "Размер"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    
    XCol = XCell.Column
    XRow = XCell.Row
    Rows("1:" & XRow - 1).Select
    Selection.Delete Shift:=xlUp
    '------создаем ключ для типа---------
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    YCol = YCell.Column
    ZCol = ZCell.Column
  
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "Сити-лайт" Or Cells(i, YCol + 1).Value = "Сити-скролл") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призма") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'-----------замена сторон--------------
    Columns(3).Select
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Б", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    '------город---------
    
Const ColtoFilter1 As Integer = 4
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 13

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("O2:O10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 15

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("i2:i6")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 18

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("G1:G4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        'конец переноса ширины столбцов
    End With
        '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 2).Value = Cells(i - 1, 2).Value And Cells(i, 3).Value = Cells(i - 1, 3).Value Then
            Rows(i).Delete
        End If
    Next i
        '-------создаем стоиомость own------

    Columns(15).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 15) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, 15).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, 11).Value = "Скролл" And Cells(i, 4).Value = "Киев" And Cells(i, 14).Value = 1 _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("Скидки").Range("AG6") _
            Else If Cells(i, 11).Value = "Скролл" And Cells(i, 4).Value = "Киев" And Cells(i, 14).Value = 2 _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("Скидки").Range("AG7") _
            Else: If Cells(i, 11).Value = "Скролл" And Cells(i, 4).Value = "Киев" And Cells(i, 14).Value = 3 _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("Скидки").Range("AG8") _
            Else: If Cells(i, 11).Value = "биллборд" And Cells(i, 4).Value = "Киев" _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("Скидки").Range("AH3") * Cells(i, 16) _
            Else: If Cells(i, 11).Value = "биллборд" And Cells(i, 4).Value = "Сумы" _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("Скидки").Range("AH4") * Cells(i, 16) _
            Else: Cells(i, 15).Value = ThisWorkbook.Worksheets("Скидки").Range("AH5") * Cells(i, 16)
    Next


    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Perekhid(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)

Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZRow, ZCol As Integer
    
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False
    
    txtCol3 = "Формат"
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    ZRow = ZCell.Row
    ZCol = ZCell.Column
    Rows("1:" & ZRow - 1).Select
    Selection.Delete Shift:=xlUp
    
    txtCol1 = "Город"
    txtCol2 = "Конструкция"


    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
  
    '------создаем ключ для типа---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "Ситискролл" Or Cells(i, YCol + 1).Value = "Ситилайт" Or _
                (Cells(i, YCol + 1).Value = "Скролл" And Cells(i, YCol + 2).Value = "1.8x1.2")) _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призма") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------переименование города-----------
    Columns(XCol).Select
    Selection.Replace What:="Кировоград (Кропивницкий )", Replacement:="Кропивницкий", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    
    '-------создаем стоиомость own------

    Columns(11).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 11) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f6").Copy
    Range(Cells(2, 11), Cells(lLastRow, 11)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '------город---------
    
Const ColtoFilter1 As Integer = 1
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 4

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("m2:m10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 6

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("g2:g6")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 15

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("e1:e4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a2")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        'конец переноса ширины столбцов
    End With
    '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Luvers(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)

Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer
    
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False
    Rows("1:4").Select
    Selection.Delete Shift:=xlUp
    
    txtCol1 = "Город"
    txtCol2 = "Конструкция"
    txtCol3 = "Формат"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
  
    '------создаем ключ для типа---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "Ситилайт" _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призма") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else If (Cells(i, YCol + 1).Value = "Скролл" And Cells(i, YCol + 2).Value = "1.2x1.8") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Скролл" And Cells(i, YCol + 2).Value = "6x3") _
            Then Cells(i, YCol).Value = "всад" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------переименование города-----------
    Columns(XCol).Select
    Selection.Replace What:="Київ", Replacement:="Киев", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    
    '-------создаем стоиомость own------

    Columns(10).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 10) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f9").Copy
    Range(Cells(2, 10), Cells(lLastRow, 10)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '------город---------
    
Const ColtoFilter1 As Integer = 1
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 4

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("p2:p10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 6

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("j2:j10")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 12

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("H1:H4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a2")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        'конец переноса ширины столбцов
    End With
        '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value = Cells(i - 1, 3).Value And Cells(i, 7).Value = Cells(i - 1, 7).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub Dovira(nameOfFile As String, nameOfFile2 As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)

Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer
    
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------открыть файлы------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2) 'Открытие файла Price
'------------определяем ключ для прайса------------

    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False

    ActiveSheet.Columns("A:BB").Hidden = False

    Cells(1, 1).EntireColumn.Insert
    ActiveSheet.AutoFilterMode = False
    lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
    Cells(1, 1).Select
    For i = lLastRow To 1 Step -1
        ActiveSheet.Cells(i, 1).Value = Cells(i, 3).Value & Cells(i, 13).Value
    Next
    Cells(1, 14).EntireColumn.Insert
    Cells(1, 14) = "Себестоимость"
    For i = lLastRow To 2 Step -1
        
                Cells(i, 14) = Application.VLookup(Cells(i, 1), Workbooks(nameOfFile2).Sheets("Price").Range( _
                                                                    Workbooks(nameOfFile2).Sheets("Price").Cells(1, 1), _
                                                                    Workbooks(nameOfFile2).Sheets("Price").Cells(lLastRow, 6)), 6, False)
    Next
    Columns(1).Delete
    
    txtCol1 = "Город"
    txtCol2 = "Тип"
    txtCol3 = "Формат"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
  
    '------создаем ключ для типа---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "Сити-лайт" _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призма") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------переименование города-----------
    Columns(XCol).Select
    Selection.Replace What:="Коломия", Replacement:="Коломыя", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        '----------преобразование сторон--------
    Columns(15).Select
    Selection.Replace What:="Б*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="А*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
       
    '------город---------
    
Const ColtoFilter1 As Integer = 2
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A200")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 12

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("N2:N10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 5

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("h2:h14")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 16

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("f1:f4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        'конец переноса ширины столбцов
    End With
    '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 11).Value = Cells(i - 1, 11).Value And Cells(i, 15).Value = Cells(i - 1, 15).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    Windows(nameOfFile2).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub RTM(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False
    txtCol1 = "Город"
    txtCol2 = "Тип"
    txtCol3 = "Размер"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
    '-------удаляем лишнюю строку и колонки-----------
    Range(Cells(XRow + 1, 16), Cells(XRow + 1, 39)).Select
    Selection.Copy
    Range(Cells(XRow, 16), Cells(XRow, 39)).Select
    Selection.PasteSpecial xlPasteAll
    Rows(XRow + 1).Select
    Rows(XRow + 1).Delete
    '-------УДАЛЯЕМ ПРОБЕЛЫ В РАЗМЕРАХ-----------
    Columns(ZCol).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------создаем ключ типа---------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1" & ":" & XRow - 1).Select
    Selection.Delete Shift:=xlUp
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "Сити-лайт" _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "скролл" And Cells(i, YCol + 2).Value = "1.86x1.3" Or _
            Cells(i, YCol + 2).Value = "1.8x1.2" Or Cells(i, YCol + 2).Value = "1.7x1.2" Or Cells(i, YCol + 2).Value = "1.86x1.27") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else If (Cells(i, YCol + 1).Value = "Щит" Or Cells(i, YCol + 1).Value = "Призма") _
            And (Cells(i, YCol + 2).Value = "3x6") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
   
    
    '-------создаем стоиомость own------

    Columns(16).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 16) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f13").Copy
    Range(Cells(2, 16), Cells(lLastRow, 16)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '----------преобразование сторон--------
    Columns(12).Select
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------город---------
    
Const ColtoFilter1 As Integer = 5
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A175")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 9
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("t2:t10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------размеры плоскостей-------------
Const ColtoFilter3 As Integer = 11
    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("n1:n10")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 18
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("l2:l4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 12).Value = Cells(i - 1, 12).Value Then
            Rows(i).Delete
        End If
    Next i
    
        '----------преобразование в числа--------
    With ActiveSheet.UsedRange.Columns(15)
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Tristar(nameOfFile As String, nameOfFile2, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)
    ActiveSheet.AutoFilterMode = False
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "Город"
    txtCol2 = "тип формат "

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    
    XCol = XCell.Column
    XRow = XCell.Row
    '------создаем ключ типа---------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Rows("1" & ":" & XRow - 1).Select
    Selection.Delete Shift:=xlUp
    
    Columns(1).Select
    Selection.Delete Shift:=xlLeft
    
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    YCol = YCell.Column

    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "призма 6х3" _
        Or Cells(i, YCol + 1).Value = "щит 6,2х3,2" _
        Or Cells(i, YCol + 1).Value = "щит 6х3" _
        Or Cells(i, YCol + 1).Value = "Щит 5,7х2,5" _
        Or Cells(i, YCol + 1).Value = "щит 5,9х 2,9") _
        Then Cells(i, YCol).Value = "биллборд" _
        Else: If Cells(i, YCol + 1).Value = "сити-лайт 1,2x1,8" _
        Then Cells(i, YCol).Value = "ситилайт" _
        Else: If Cells(i, YCol + 1).Value = "скролл 3,14х2,32" _
        Then Cells(i, YCol).Value = "скролл" _
        Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------создаем стоиомость own------

    Columns(17).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 17) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f18").Copy
    Range(Cells(2, 17), Cells(lLastRow, 17)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '----------преобразование сторон--------
    Columns(15).Select
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="А*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="В*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
            Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.FormulaR1C1 = "A"
    
 '---------добавляем GRP--------------
    Cells(1, 16).EntireColumn.Insert
    Cells(1, 16) = "GRP"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        Cells(i, 16) = Application.IfError(Application.VLookup(Cells(i, 10), Workbooks(nameOfFile2).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile2).Sheets("GRP").Cells(1, 10), _
                                                                    Workbooks(nameOfFile2).Sheets("GRP").Cells(lLastRow, 13)), 4, False), "-")
    Next

    '------город---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A175")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 3
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("y2:y10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 20
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("q2:q4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 13).Value = Cells(i - 1, 13).Value And Cells(i, 15).Value = Cells(i - 1, 15).Value Then
            Rows(i).Delete
        End If
    Next i
    

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    Windows(nameOfFile2).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub Sean(nameOfFile As String, nameOfFile2, nameOfSheet, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Add
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)
    ActiveSheet.AutoFilterMode = False
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False
'-----------ID Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "ID"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ID Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Код Доорс Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Код Doors"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Код Доорс Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("B1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Город Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Город"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Город Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("C1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Район Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Район"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Район Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("D1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Фото 1 Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Фото 1"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Фото 1 Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("E1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Карта Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Карта"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Карта Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("F1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Адрес Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Адрес"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Адрес Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("H1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Тип носителя Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Тип носителя"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Тип носителя Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("J1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------OTS Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "OTS"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "OTS Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("N1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------GRP Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "GRP"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "GRP Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("O1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------Свет Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Свет"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Свет Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("P1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------Описание Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Описание"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Описание Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("Q1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------Цена Прайс и Занятость Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Цена Прайс"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Цена Прайс Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol + 22) & lLastRow).Copy
    Workbooks("Книга1").Activate
    ActiveWorkbook.ActiveSheet.Range("S1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------замена сторон--------------
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("L:L").Select
    ActiveSheet.Paste
    Columns("L:L").Select
    Selection.Replace What:="* А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="* В", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="* С", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'---------замена типов щитов-----------------------

    Workbooks("Книга1").Activate
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Columns("I:I").Select
    Dim fndList As Variant
    Dim x As Long
    fndList = Array("Щит высокий 6х3м", "Призма высокая 6х3м", "Гусь-призма 6х3м", "Призма VIP 6х3м", "Призма 6х3м", "Гусь 6х3м", "Щит 6х3м")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:="биллборд", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
'-----------вносим себестоимость---------------
    Workbooks("Книга1").Activate
    Columns("J:J").Select
    Selection.Copy
    Columns("R:R").Select
    ActiveSheet.Paste
    Dim Rng As Range
    Dim InputRng As Range, ReplaceRng As Range
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set InputRng = ActiveWorkbook.ActiveSheet.Range("R1:" & "R" & lLastRow)
    Set ReplaceRng = ThisWorkbook.Sheets("Скидки").Range("I2:j8")
    For Each Rng In ReplaceRng.Columns(1).Cells
        InputRng.Replace What:=Rng.Value, Replacement:=Rng.Offset(0, 1).Value
    Next
    
 '-----------Вставить названия столбцов из сетки City и TYPE-----------
    Workbooks("Книга1").Activate
    Range("G1").Value = "Сайт"
    Range("I1").Value = "TYPE"
    Range("K1").Value = "Размер"
    Range("L1").Value = "Сторона"
    Range("M1").Value = "ДХ"
    Range("R1").Value = "Себестоимость"
    Range("A1").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Workbooks.Add
'-----------ID City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "ID"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ID City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Код Doors City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Код Doors"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Код Doors City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("B1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Город City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Город"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Город City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("C1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Район City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Район"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Район City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("D1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Сайт City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Сайт"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Сайт City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("G1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Адрес City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Адрес"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Адрес City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("H1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Тип City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Тип"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Тип City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("J1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------Размер City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Размер"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Размер City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("K1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Сторона City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Сторона"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Сторона City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("L1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------ДХ City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "ДХ"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ДХ City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("M1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------OTS City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "OTS"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "OTS City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("N1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------GRP City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "GRP"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "GRP City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("O1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Свет City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Свет"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Свет City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("P1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Описание City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Описание"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Описание City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("Q1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------Прайс и Занятость City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Прайс"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Прайс City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("Книга2").Activate
    ActiveWorkbook.ActiveSheet.Range("S1").PasteSpecial Paste:=xlPasteAll
'-----------занятость---------------
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol + 2) & XRow + 1 & ":" & ReturnName(XCol + 23) & lLastRow).Copy
    Workbooks("Книга2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range("T1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------обработка сторон--------------
    Columns("L:L").Select
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'------переименование скролл и сити в ситилайт-----------
    Workbooks("Книга2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 1 Step -1
        If Cells(i, 10).Value = "Сити-лайт" _
        Or (Cells(i, 10).Value = "Скролл" And Cells(i, 11).Value = "1.8x1.2") _
        Then Cells(i, 9).Value = "ситилайт" _
        Else If (Cells(i, 10).Value = "Скролл" And Cells(i, 11).Value = "3x1.5") _
        Then Cells(i, 9).Value = "щит" _
        Else: If (Cells(i, 10).Value = "Призма" And Cells(i, 11).Value = "3x6") _
        Then Cells(i, 9).Value = "биллборд" _
        Else: Cells(i, 9).Value = Cells(i, 10)
    Next
'------внесение себестоимости-----------
    Workbooks("Книга2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 1 Step -1
        If (Cells(i, 9).Value = "ситилайт" And Cells(i, 8).Value = "*Дерибасовск*") _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("Скидки").Range("J10") _
        Else: If (Cells(i, 9).Value = "ситилайт" And Cells(i, 8).Value = "*Аркадия*") _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("Скидки").Range("J11") _
        Else: If Cells(i, 9).Value = "Скролл" _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("Скидки").Range("J13") _
        Else: If Cells(i, 9).Value = "биллборд" _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("Скидки").Range("J14") _
        Else: Cells(i, 18).Value = ThisWorkbook.Sheets("Скидки").Range("J12")
    Next
'-----------соединяем сетки-------------
    Workbooks("Книга2").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Workbooks("Книга1").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
'----------преобразование в числа--------
    With ActiveSheet.UsedRange.Columns(15)
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With

    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 9
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("U2:U10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 21
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("M2:M4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues
        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
            '-------------------удалить дубликаты--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 12).Value = Cells(i - 1, 12).Value Then
            Rows(i).Delete
        End If
    Next i
    '-------------проставляем город------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(2, 3).Select
    Cells(2, 3).Value = "Одесса"
    Cells(2, 3).Select
    Selection.AutoFill Destination:=Range(Cells(2, 3), Cells(lLastRow, 3)), Type:=xlFillDefault

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    Windows(nameOfFile2).Close
    Windows("Книга1").Close
    Windows("Книга2").Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub Mallis(nameOfFile As String, nameOfFile1 As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, rngReserv, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, VlLastRow, lLastCol As Integer
Dim YCell As Object
Dim YRow, YCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  'Открытие файла
    ActiveSheet.AutoFilterMode = False

    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False
    txtCol2 = "НОСИТЕЛЬ"

    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------создаем ключ типа---------
    Rows(1 & ":" & YRow - 1).Select
    Selection.Delete Shift:=xlUp
    Rows(YRow & ":" & YRow + 1).Select
    Selection.Delete Shift:=xlUp
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "скрол" _
            Then Cells(i, YCol).Value = "скролл" _
            Else: If (Cells(i, YCol + 1).Value = "призма" Or Cells(i, YCol + 1).Value = "щит") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'------внесение себестоимости-----------
    Columns(YCol + 1).Select
    Selection.Insert Shift:=xlToRight
    Cells(1, YCol + 1) = "Себестоимость"
    For i = lLastRow To 1 Step -1
        If Cells(i, YCol + 2).Value = "щит" _
        Then Cells(i, YCol + 1).Value = ThisWorkbook.Sheets("Скидки").Range("M2") _
        Else: If Cells(i, YCol + 2).Value = "призма" _
        Then Cells(i, YCol + 1).Value = ThisWorkbook.Sheets("Скидки").Range("M3") _
        Else: If Cells(i, YCol + 2).Value = "скрол" _
        Then Cells(i, YCol + 1).Value = ThisWorkbook.Sheets("Скидки").Range("M4")
    Next
    '---------удалить пустые строки----------
    Dim r As Long
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For r = LastRow To 2 Step -1
    If Application.CountA(Rows(r)) = 0 Then
        Rows(r).Delete
    End If
    Next r
    '-------Вставляем город------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Columns(3).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 3) = "Город"
    Cells(2, 3) = "Киев"
    Cells(2, 3).Copy
    Range(Cells(3, 3), Cells(lLastRow, 3)).Select
    Selection.PasteSpecial Paste:=xlAll
'-----------замена сторон--------------
    Columns(5).Select
    Selection.Replace What:="разд-ль", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="В", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'----------- убираем стороны из адресов--------------
    Columns(4).Select
    Selection.Replace What:="(А5)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(А4)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(А3)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(А2)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(А1)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(А)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(В5)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(В4)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(В3)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(В2)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(В1)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(В)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------добавляем GRP--------------
    Windows(nameOfFile1).Activate
    VlLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Windows(nameOfFile).Activate
    lLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Cells(1, 14).EntireColumn.Insert
    Cells(1, 14) = "GRP"
    For i = lLastRow To 2 Step -1
        Cells(i, 14) = Application.IfError(Application.VLookup(Cells(i, 1), Workbooks(nameOfFile1).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile1).Sheets("GRP").Cells(1, 1), _
                                                                    Workbooks(nameOfFile1).Sheets("GRP").Cells(VlLastRow, 2)), 2, False), "")
    Next
    Windows(nameOfFile1).Close

    '------город---------
    
Const ColtoFilter1 As Integer = 3
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 6
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("r2:r10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 15
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("j2:j4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
'Set StartCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Range(Cells(XRow, 1))
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 5).Value = Cells(i - 1, 5).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub Alhor(nameOfFile As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, rngReserv, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, VlLastRow, lLastCol As Integer
Dim YCell As Object
Dim YRow, YCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Columns("A:BB").Hidden = False 'расскрываем все столбцы
    ActiveWindow.FreezePanes = False 'убрать закрепление областей
    txtCol2 = "тип конструкции"

    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------создаем ключ типа---------
    Rows(1 & ":" & YRow - 1).Select
    Selection.Delete Shift:=xlUp
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "скролл 2,30х3,140" _
            Then Cells(i, YCol).Value = "скролл" _
            Else: If (Cells(i, YCol + 1).Value = "призма 3х6" Or Cells(i, YCol + 1).Value = "щит 3х6,2" Or _
                Cells(i, YCol + 1).Value = "щит 3х6" Or Cells(i, YCol + 1).Value = "щит 3,2х6,2") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else: If (Cells(i, YCol + 1).Value = "сити-скроллер 1.2х1.8" Or Cells(i, YCol + 1).Value = "сити-лайт 1,2Х1,8") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'------внесение себестоимости-----------
    Columns(11).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 11) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f10").Copy
    Range(Cells(2, 11), Cells(lLastRow, 11)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '-------Вставляем город------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Columns(2).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 2) = "Город"
    Cells(2, 2) = "Киев"
    Cells(2, 2).Copy
    Range(Cells(3, 2), Cells(lLastRow, 2)).Select
    Selection.PasteSpecial Paste:=xlAll
'-----------замена сторон--------------
    Columns(4).Select
    Selection.Replace What:="А*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Б*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------город---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 5
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("Q2:Q10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

    '--------------размер------------------
Const ColtoFilter3 As Integer = 6
    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("Форматы").Range("K2:K10")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 16
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("j2:j4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("A1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по размеру----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value = Cells(i - 1, 3).Value And Cells(i, 4).Value = Cells(i - 1, 4).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub CityDnepr(nameOfFile As String, nameOfSheet As String, nameOfSheetBoard As String, nameOfSheetCity As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Sheets.Add
    Workbooks(nameOfFile).Sheets.Add
    
'-----------переносим иформацию по щитам Адрес-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetBoard).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Адрес"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "АдресЩитаСитиДнепр, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-------------вставляем тип плоскости------------
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(lLastCol).Copy
    Columns(ReturnName(lLastCol + 1) & ":" & ReturnName(lLastCol + 2)).PasteSpecial Paste:=xlPasteAll
    Cells(1, lLastCol + 1).Value = "Тип плоскости"
    Cells(2, lLastCol + 1).Value = "биллборд"
    Cells(1, lLastCol + 2).Value = "Размер"
    Cells(2, lLastCol + 2).Value = "6х3"
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & 2).Copy
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & lLastRow).PasteSpecial Paste:=xlPasteValues
'-----------переносим иформацию по щитам Фото и до конца-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetBoard).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Фото"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ФотоЩитаСитиДнепр, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(XCol) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(lLastCol + 1) & 1).PasteSpecial Paste:=xlPasteAll
    End If
'-----------Удалить пустые столбцы по первой строке---------
    For i = 30 To 1 Step -1
        If Cells(1, i).Value = 0 Then
            Columns(i).Delete
            i = i - 1
        End If
    Next i

'-----------переносим иформацию по ситилайтам Адрес-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Адрес"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "АдресСитилайтаСитиДнепр, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист2").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-------------вставляем тип плоскости------------
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(lLastCol).Copy
    Columns(ReturnName(lLastCol + 1) & ":" & ReturnName(lLastCol + 2)).PasteSpecial Paste:=xlPasteAll
    Cells(1, lLastCol + 1).Value = "Тип плоскости"
    Cells(2, lLastCol + 1).Value = "ситилайт"
    Cells(1, lLastCol + 2).Value = "Размер"
    Cells(2, lLastCol + 2).Value = "1.2x1.8"
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & 2).Copy
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & lLastRow).PasteSpecial Paste:=xlPasteValues

'-----------переносим иформацию по ситилайтам Фото и до конца-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Фото"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ФотоСитилайтСитиДнепр, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(XCol) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист2").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(lLastCol + 1) & 1).PasteSpecial Paste:=xlPasteAll
    End If
'-----------Удалить пустые столбцы по первой строке---------
    For i = 30 To 1 Step -1
        If Cells(1, i).Value = 0 Then
            Columns(i).Delete
            i = i - 1
        End If
    Next i
'-----------соединяем сетки-------------
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Workbooks(nameOfFile).Sheets("Лист2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    '-------создаем стоиомость own------

    Columns(9).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 9) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f17").Copy
    Range(Cells(2, 9), Cells(lLastRow, 9)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    
'----------убрать пробелы--------
    With ActiveSheet.UsedRange.Columns(13)
        .Replace " ", ""
    End With

'----------преобразование в числа--------
    With ActiveSheet.UsedRange.Columns(13)
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 4
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("x2:x10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 15

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=1
        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    '-------------проставляем город------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(2, 1).Select
    Cells(2, 1).Value = "Днепр"
    Cells(2, 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow, 1)), Type:=xlFillDefault

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Prospect(nameOfFile As String, nameOfFile1 As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, rngReserv, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, VlLastRow, lLastCol As Integer
Dim YCell As Object
Dim YRow, YCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  'Открытие файла
    ActiveSheet.AutoFilterMode = False

    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False 'убрать закрепление областей
    Cells.MergeCells = False 'убрать объединение ячеек

    txtCol2 = "Вид"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------создаем ключ типа---------
    Rows(1 & ":" & YRow - 1).Select
    Selection.Delete Shift:=xlUp
    Rows(YRow & ":" & YRow + 1).Select
    Selection.Delete Shift:=xlUp
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 4 Step -1
        If Cells(i, YCol + 1).Value = "Скролл 3,14х2,32" _
            Then Cells(i, YCol).Value = "скролл" _
            Else: If (Cells(i, YCol + 1).Value = "Щит 3х6" Or Cells(i, YCol + 1).Value = "Призма3х6" Or Cells(i, YCol + 1).Value = "Щит 3,2х6,2") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'------внесение себестоимости-----------
    Columns(12).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 12) = "Себестоимость"
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f19").Copy
    Range(Cells(4, 12), Cells(lLastRow, 12)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
'-----------замена сторон--------------
    Columns(7).Select
    Selection.Replace What:="р/п", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="В*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="А*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------добавляем GRP--------------
    Windows(nameOfFile1).Activate
    VlLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Windows(nameOfFile).Activate
    lLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Cells(1, 11).EntireColumn.Insert
    Cells(1, 11) = "GRP"
    For i = lLastRow To 2 Step -1
        Cells(i, 11) = Application.IfError(Application.VLookup(Cells(i, 3), Workbooks(nameOfFile1).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile1).Sheets("GRP").Cells(1, 2), _
                                                                    Workbooks(nameOfFile1).Sheets("GRP").Cells(VlLastRow, 3)), 2, False), "")
    Next
    Windows(nameOfFile1).Close
    '-------удалить пустые строки-----------
    For i = 5 To 1 Step -1
        If Cells(i, 1).Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 4
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("z2:z10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 15
Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

                                                           
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=1

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    '-------удалить дубликаты скроллов------------
    Columns(6).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Selection.Replace What:="-*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 6).Value = Cells(i - 1, 6).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i
    Columns(6).Delete
    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close
Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Megapolis(nameOfFile As String, nameOfFile1 As String, nameOfFile2 As String, nameOfSheet As String, nameOfSheetBoard As String, nameOfSheetCity As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol, ZRow As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Sheets.Add
    Workbooks(nameOfFile).Sheets.Add
    
'-----------переносим иформацию по щитам-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetBoard).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row

    txtCol = "Город"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Город_щит_Мега_Харьков, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    txtCol = ""
    Set ZCell = Workbooks(nameOfFile).ActiveSheet.Range(ReturnName(1) & XRow & ":" & ReturnName(1) & lLastRow).Find(txtCol)
    ZCol = ZCell.Column
    ZRow = ZCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(lLastCol) & ZRow - 1).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------переносим иформацию по ситилайтам-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Город"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Город_сити_Мега_Харьков, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
'-----------вставляем столбец район(для объединения с Днепром)-------
    Columns(9).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, 9) = "Район"

'-----------переносим иформацию по Днепру-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  'Открытие файла
    Workbooks(nameOfFile1).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Город"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Город_щит_Мега_Днепр, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
    '------создаем ключ типа---------
    txtCol = "Рaзмер"  ' метка для столбца
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Размер_Мега_Харьков, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Cells(lLastRow, YCol - 1).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "3х6" Or Cells(i, YCol + 1).Value = "2,9х5,9") _
            And (Cells(i, YCol + 2).Value = "призма" _
            Or Cells(i, YCol + 2).Value = "флаг" _
            Or Cells(i, YCol + 2).Value = "стандарт" _
            Or Cells(i, YCol + 2).Value = "чебуршка-призма" _
            Or Cells(i, YCol + 2).Value = "гусь") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else If (Cells(i, YCol + 1).Value = "3,1х2,2" Or Cells(i, YCol + 1).Value = "3,33х2,3") _
            And (Cells(i, YCol + 2).Value = "скролл" Or Cells(i, YCol + 2).Value = "бэклайт") _
            Then Cells(i, YCol).Value = "скролл" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 2)
    Next
    End If
'------внесение себестоимости-----------
    Columns(12).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 12) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol - 1).Value = "*VIP*") _
        Then Cells(i, 12).Value = Cells(i, 13) * ThisWorkbook.Sheets("Скидки").Range("q3") _
        Else: Cells(i, 12).Value = Cells(i, 13) * ThisWorkbook.Sheets("Скидки").Range("q2")
    Next
'-----------замена сторон--------------
    Columns(8).Select
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Б", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------добавляем GRP--------------
     Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)  'Открытие файла
    Windows(nameOfFile2).Activate
    VlLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Windows(nameOfFile).Activate
    lLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Cells(1, 11).EntireColumn.Insert
    Cells(1, 11) = "GRP"
    For i = lLastRow To 2 Step -1
        Cells(i, 11) = Application.IfError(Application.VLookup(Cells(i, 3), Workbooks(nameOfFile2).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile2).Sheets("GRP").Cells(1, 2), _
                                                                    Workbooks(nameOfFile2).Sheets("GRP").Cells(VlLastRow, 3)), 2, False), "")
    Next
    Windows(nameOfFile2).Close
'------фильтр по занятости------------------
    colorReserv = RGB(204, 153, 255)
    colorFree = RGB(255, 255, 255)
    For i = lLastRow To 2 Step -1
            If Cells(i, 15).Interior.Color = colorReserv _
            Then Cells(i, 15).Value = "Резерв" _
            Else If Cells(i, 15).Interior.Color = colorFree _
            Then Cells(i, 15).Value = "Свободно"
    Next i
    
    Columns(1).Select
    Selection.Replace What:="Хaрьков", Replacement:="Харьков", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------город---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 5
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("s2:s10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 15
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("k2:k4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)


Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues


        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    Windows(nameOfFile1).Close
    wb.Close
Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub Bomond(nameOfFile As String, nameOfFile1 As String, nameOfSheet1 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, rngReserv, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, VlLastRow, lLastCol As Integer
Dim YCell As Object
Dim YRow, YCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet1).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False 'убрать закрепление областей
    Cells.MergeCells = False 'убрать объединение ячеек

    txtCol = "Ценовая категория:"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------создаем ключ типа---------
    Columns(YCol).Select
    Application.CutCopyMode = False
    Selection.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Columns(YCol).Select
    '-----------замена типов---------------
    Dim fndList As Variant
    Dim x As Long
    fndList = Array("щит низкий", "щит высокий", "гусак", "чебурашка", "призма VIP", "призма")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:="биллборд", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    fndList = Array("скролл VIP", "скролл", "")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:="скролл", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
'-----------вносим себестоимость---------------
    Columns(YCol + 1).Select
    Selection.Copy
    Selection.Insert Shift:=xlToLeft

    Dim Rng As Range
    Dim InputRng As Range, ReplaceRng As Range
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set InputRng = ActiveWorkbook.ActiveSheet.Range(Cells(1, YCol + 2), Cells(lLastRow, YCol + 2))
    Set ReplaceRng = ThisWorkbook.Sheets("Скидки").Range("S3:T14")
    For Each Rng In ReplaceRng.Columns(1).Cells
        InputRng.Replace What:=Rng.Value, Replacement:=Rng.Offset(0, 1).Value
    Next
    Cells(1, YCol + 2).Select
    Cells(1, YCol + 2) = "Себестоимость"
'-----------замена сторон--------------
    txtCol = "Сторона:"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    
    YCol = YCell.Column
    YRow = YCell.Row

    Columns(YCol).Select
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Б", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------добавляем GRP--------------
    'преобразовать коды дорс в числа
    
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ThisWorkbook.Sheets("Скидки").Range("B15").Copy
    Range("B2:" & "B" & lLastRow).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    'добавляем значения GRP
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  'Открытие файла
    Windows(nameOfFile1).Activate
    VlLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Windows(nameOfFile).Activate
    lLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Cells(1, 11).EntireColumn.Insert
    Cells(1, 11) = "GRP"
    For i = lLastRow To 2 Step -1
        Cells(i, 11) = Application.IfError(Application.VLookup(Cells(i, 2), Workbooks(nameOfFile1).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile1).Sheets("GRP").Cells(1, 2), _
                                                                    Workbooks(nameOfFile1).Sheets("GRP").Cells(VlLastRow, 3)), 2, False), "")
    Next
    Windows(nameOfFile1).Close
    
    '-------------проставляем город------------
    Columns(2).Select
    Selection.Insert Shift:=xlRight
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(2, 2).Select
    Cells(2, 2).Value = "Одесса"
    Cells(2, 2).Select
    Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lLastRow, 2)), Type:=xlFillDefault
    '-------------определяем занятость------------
    txtCol = "По:"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Rows(1).Cells.Find(txtCol)
    
    YCol = YCell.Column
    YRow = YCell.Row

    Columns(YCol).Select
    Selection.Insert Shift:=xlToLeft
    Workbooks(nameOfGeneralFile).Worksheets("Скидки").Range("g15").Copy
    Cells(1, YCol).Select
    Selection.PasteSpecial Paste:=xlPasteAll
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "" Then Cells(i, YCol + 1).Value = Cells(1, YCol).Value + 365
            If Cells(i, YCol - 1).Value <= Cells(1, YCol).Value And Cells(1, YCol).Value <= Cells(i, YCol + 1).Value _
                Then Cells(i, YCol) = "Свободна" _
                Else Cells(i, YCol) = "Занята"
        
    Next
    '------город---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 12
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("v2:v10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 6
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("N2:N4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet1).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close
Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub

Sub ThreeSixDnepr(nameOfFile As String, nameOfSheet As String, nameOfSheetRegion As String, nameOfSheetCity As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Sheets.Add
    
'-----------переносим иформацию по Региону-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetRegion).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "місто"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Область_3x6Dnepr, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
    Columns(XCol + 1).Select
    Selection.Insert Shift:=xlToRight
    Cells(1, XCol + 1).Value = "Район"

'-----------переносим иформацию по городу-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "Місто"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Город_3x6Dnepr, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
    '--------замена размеров---------
    txtCol = "розмір"  ' метка для столбца
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "размер_Область3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Columns(YCol).Select
    Selection.Replace What:="х", Replacement:="x", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
    '------создаем ключ типа---------
    txtCol = "формат"  ' метка для столбца
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "формат_Область3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Cells(lLastRow, YCol - 1).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 2).Value = "2,85x5,95" _
            Or Cells(i, YCol + 2).Value = "3x6" _
            Or Cells(i, YCol + 2).Value = "2,90x5,90" _
            Or Cells(i, YCol + 2).Value = "2,95x6,0" _
            Or Cells(i, YCol + 2).Value = "3,0x5,9" _
            Or Cells(i, YCol + 2).Value = "2,85x5,9" _
            Or Cells(i, YCol + 2).Value = "2,85x5,90" _
            Or Cells(i, YCol + 2).Value = "2,95x5,90" _
            And Cells(i, YCol + 1).Value = "billboard") _
            Then Cells(i, YCol).Value = "биллборд" _
            Else If (Cells(i, YCol + 2).Value = "1,8x1,2" And Cells(i, YCol + 1).Value = "sity-light") _
            Then Cells(i, YCol).Value = "ситилайт" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 2)
    Next
    End If
'------внесение себестоимости-----------
    txtCol = "ст."  ' метка для столбца
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "цена_Область3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Columns(YCol + 1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, YCol + 1) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("f20").Copy
    Range(Cells(2, YCol + 1), Cells(lLastRow, YCol + 1)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    End If
'-----------замена сторон--------------
    txtCol = "ст."  ' метка для столбца
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "сторона_Область3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Range(Cells(2, YCol), Cells(lLastRow, YCol)).Select
    Selection.Replace What:="А*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Б*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="С*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
'------фильтр по занятости------------------
   
    Columns(1).Select
    Dim fndList, fndCity As Variant
    Dim x As Long
    fndList = Array("Дніпро", "Павлоград ", "Новомосковськ", "Кам'янське", "Синельникове", "Нікополь", "Дрогобич")
    fndCity = Array("Днепр", "Павлоград", "Новомосковск", "Каменское", "Синельниково", "Никополь", "Дрогобыч")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:=fndCity(x), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    '------город---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 3
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("AA2:AA10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 17
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("s2:s4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)


Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues


        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 12).Value = Cells(i - 1, 12).Value And Cells(i, 14).Value = Cells(i - 1, 14).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close
Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub NashaSprava(nameOfFile As String, nameOfSheet As String, nameOfSheetCity As String, nameOfFile2 As String, nameOfFile3 As String, nameOfFile4 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Sheets.Add
    
'-----------переносим иформацию по щитам-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    Workbooks(nameOfFile).Sheets.Add
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row

    txtCol = "link"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Link_Щит_НашаСправа, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Columns(6).Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlLeft
    Cells(XRow, 6).Value = "Type"
    For i = lLastRow To XRow + 1 Step -1
        If InStr(1, Cells(i, 4), "призма") <> 0 _
            Then Cells(i, 6).Value = "призма" _
            Else: Cells(i, 6).Value = "биллборд"
    Next
    Range(ReturnName(1) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
    '-----------переносим иформацию по ситилайтам-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)  'Открытие файла
    Workbooks(nameOfFile2).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
            '---меняем местами столбцы код и код дорс
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 4).End(xlUp).Row
    txtCol = "link"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Link_Ситилайт_НашаСправа, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Columns(6).Select
    Selection.Insert Shift:=xlRight
    Cells(XRow, 6).Value = "Type"
    For i = lLastRow To XRow + 1 Step -1
        If InStr(1, Cells(i, 4), "скролл") <> 0 _
            Then Cells(i, 6).Value = "ситискролл" _
            Else: Cells(i, 6).Value = "ситилайт"
    Next
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
    '-----------переносим иформацию по скроллам-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile3)  'Открытие файла
    Workbooks(nameOfFile3).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    txtCol = "link"  ' метка для столбца
    Set XCell = Workbooks(nameOfFile3).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Link_Скролл_НашаСправа, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Columns(6).Select
    Selection.Insert Shift:=xlRight
    Cells(XRow, 6).Value = "Type"
    Cells(XRow + 1, 6).Value = "скролл"
    Cells(XRow + 1, 6).Select
    Selection.AutoFill Destination:=Range(Cells(XRow + 1, 6), Cells(lLastRow, 6)), Type:=xlFillDefault
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("Лист1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If

'------внесение себестоимости-----------
    Columns(6).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Dim Rng As Range
    Dim InputRng As Range, ReplaceRng As Range
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    Set InputRng = ActiveWorkbook.ActiveSheet.Range(Cells(1, 6), Cells(lLastRow, 6))
    Set ReplaceRng = ThisWorkbook.Sheets("Скидки").Range("V3:W7")
    For Each Rng In ReplaceRng.Columns(1).Cells
        InputRng.Replace What:=Rng.Value, Replacement:=Rng.Offset(0, 1).Value
    Next
    Cells(1, 6).Select
    Cells(1, 6) = "Себестоимость"
    Range(ReturnName(6) & 2 & ":" & ReturnName(6) & lLastRow).Select
    Selection.NumberFormat = "0.00"
'-----------замена типов--------------
    Columns(7).Select
    Selection.Replace What:="призма", Replacement:="биллборд", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="ситискролл", Replacement:="ситилайт", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'-----------замена сторон--------------
    txtCol = "Сторона"  ' метка для столбца
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "сторона_НашаСправа, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Range(Cells(2, YCol), Cells(lLastRow, YCol)).Select
    Selection.Replace What:="розд.", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="В", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
'-------------проставляем город------------
    Columns(2).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(2, 2).Select
    Cells(2, 2).Value = "Львов"
    Cells(2, 2).Select
    Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lLastRow, 2)), Type:=xlFillDefault
 '---------добавляем GRP--------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile4)  'Открытие файла
    Windows(nameOfFile4).Activate
    VlLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Windows(nameOfFile).Activate
    lLastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Cells(1, 8).EntireColumn.Insert
    Cells(1, 8) = "GRP"
    For i = lLastRow To 2 Step -1
        Cells(i, 8) = Application.IfError(Application.VLookup(Cells(i, 3), Workbooks(nameOfFile4).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile4).Sheets("GRP").Cells(1, 2), _
                                                                    Workbooks(nameOfFile4).Sheets("GRP").Cells(VlLastRow, 3)), 2, False), "")
    Next
    Windows(nameOfFile4).Close
    
    '------город---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 9
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("AB2:AB10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 11
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("T2:T4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)
'--------исключаем села и регионы-----------------
Const ColtoFilter5 As Integer = 5
    
Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:="<>*продан*", Operator:=xlFilterValues
        '------------фильтр исключения сел и регионов----------------
        .AutoFilter Field:=ColtoFilter5, Criteria1:="<>*с.*", Operator:=xlAnd, Criteria2:="<>*м.*"


        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        '.Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        ws2.Cells(2, 1).PasteSpecial Paste:=xlPasteAll 'xlPasteColumnWidths'конец переноса ширины столбцов
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
    End With
    
    '-------------------удалить дубликаты--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value <> 0 Then _
            If Cells(i, 3).Value = Cells(i - 1, 3).Value _
            Then Rows(i).Delete
            
    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    Windows(nameOfFile2).Close
    Windows(nameOfFile3).Close
    wb.Close
Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub MegapolisUA(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol, ZRow As Integer

Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet2).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "Город"
    txtCol2 = "Тип"

    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol1)
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    
    '------замена формата---------
    Rows("1:" & (XRow - 1)).Select
    Selection.Delete Shift:=xlUp
    
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    Columns(YCol).Select
    Selection.Replace What:="стандарт", Replacement:="биллборд", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '-----------замена сторон--------------
    Columns(YCol + 1).Select
    Selection.Replace What:="А", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Б", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'------внесение себестоимости-----------
    Columns(12).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 12) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        Cells(i, 12).Value = Cells(i, 11) * (1 - ThisWorkbook.Sheets("Скидки").Range("B22"))
    Next

    '------фильтр по занятости------------------
    colorReserv = RGB(204, 153, 255)
    colorFree = RGB(255, 255, 255)
    For i = lLastRow To 2 Step -1
            If Cells(i, 13).Interior.Color = colorReserv _
            Then Cells(i, 13).Value = "Резерв" _
            Else If Cells(i, 13).Interior.Color = colorFree _
            Then Cells(i, 13).Value = "Свободно"
    Next i
    '------добавляем ключ по гео------------------
    Columns(1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    For i = lLastRow To 2 Step -1
            Cells(i, 1).Value = Cells(i, 2).Value & Cells(i, 4).Value
    Next i
    
    '------город---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 8
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("s2:s10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 14
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("k2:k4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)


Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues


        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
    
    '-------------------удалить дубликаты--------------------
'    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
'    For i = lLastRow To 2 Step -1
'        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
'            Rows(i).Delete
'        End If
'    Next i

    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close
Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Sub T52(nameOfFile As String, nameOfSheet1 As String, nameOfSheet2 As String, pathDir As String, nameOfGeneralFile As String)
Dim rngFree, rngCity, rngType, rngSize, startCell As Range
Dim Flag As Boolean
Dim ws As Worksheet
Dim ws2
Dim lLastRow, lLastCol As Integer
Dim XCell, YCell, ZCell As Object
Dim XCol, XRow, YCol, ZCol As Integer

'---------убираем старые данные-----------
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet2).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------город------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  'Открытие файла
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "Область"
    
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol1)
    
    XCol = XCell.Column
    XRow = XCell.Row
    
    '------Формат---------
    Rows("1:" & (XRow - 1)).Select
    Selection.Delete Shift:=xlUp
   
    Columns(XCol + 2).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, XCol + 2).Select
    Cells(1, XCol + 2) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, XCol + 2).Select
    For i = lLastRow To 2 Step -1
        Cells(i, XCol + 2).Value = "биллборд"
    Next
    
    '------внесение себестоимости-----------
    Columns(9).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 10) = "Себестоимость"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        Cells(i, 10).Value = Cells(i, 9) * (1 - ThisWorkbook.Sheets("Скидки").Range("B23"))
    Next
    '------добавляем ключ по гео------------------
    Columns(1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    For i = lLastRow To 2 Step -1
            Cells(i, 1).Value = Cells(i, 4).Value & Cells(i, 5).Value
    Next i
        '--------перевод с укр. на рус.------------
    Columns(1).Select
    Dim RuName, UAName As Variant
    Dim x As Long
    UAName = Array("ЧеркаськаСміла", "ЧеркаськаКанів", "ЧеркаськаКорсунь-Шевченківський ", "ЧеркаськаЗвенигородка", "Кіровоградська Світловодськ", _
        "Кіровоградська Бобринець", "Кіровоградська Олександрія", "Кіровоградська Знам'янка", "Полтавська Горішні Плавні ", "Полтавська Кобеляки", "Полтавська Миргород", _
        "Дніпропетровська Синельникове", "Дніпропетровська Павлоград", "Дніпропетровська Вільногірськ", "Полтавська Лубни", "ДонецькаВолноваха", "Кіровоградська Олександрівка", _
        "Дніпропетровська Апостолово", "Дніпропетровська П'ятихатки", "Дніпропетровська Жовті Води ", "Полтавська Кременчук", "ДонецькаСлов'янськ", "ДонецькаКостянтинівка", _
        "ДонецькаМирноград", "ЧеркаськаВатутіно", "ЧеркаськаУмань", "ЗапорізькаПологи", "Дніпропетровська Нікополь", "ДонецькаДружківка", "Київська Тараща", "Київська Боярка", _
        "Полтавська Карлівка", "ЗапорізькаВасилівка", "Київська Бориспіль", "ДонецькаБахмут ", "СумськаРомни", "ХмельницькаСтарокостянтинів", _
        "ЗапорізькаДніпрорудне", "ХмельницькаКам'янець-Подільський", "СумськаШостка", "ДонецькаПокровськ", _
        "Дніпропетровська Софієвка", "ЛуганськаРубіжне", "ДонецькаТорецьк", "ДонецькаКраматорськ", "Полтавська Хорол", _
        "ЗапорізькаБердянськ", "ЛуганськаКремінна")
    RuName = Array("ЧеркасскаяСмела", "ЧеркасскаяКанев", "ЧеркасскаяКорсунь-Шевченковский", "ЧеркасскаяЗвенигородка", "КировоградскаяСветловодск", _
        "КировоградскаяБобринец", "КировоградскаяАлександрия", "КировоградскаяЗнамянка", "ПолтавскаяГоришни Плавни", "ПолтавскаяКобеляки", "ПолтавскаяМиргород", _
        "ДнепропетровскаяСинельниково", "ДнепропетровскаяПавлоград", "ДнепропетровскаяВольногорск", "ПолтавскаяЛубны", "ДонецкаяВолноваха", "КировоградскаяАлександровка", _
        "ДнепропетровскаяАпостолово", "ДнепропетровскаяПятихатки", "ДнепропетровскаяЖелтые Воды", "ПолтавскаяКременчуг", "ДонецкаяСлавянск", "ДонецкаяКонстантиновка", _
        "ДонецкаяМирноград", "ЧеркасскаяВатутино", "ЧеркасскаяУмань", "ЗапорожскаяПологи", "ДнепропетровскаяНикополь", "ДонецкаяДружковка", "КиевскаяТараща", "КиевскаяБоярка", _
        "ПолтавскаяКарловка", "ЗапорожскаяВасильевка", "КиевскаяБорисполь", "ДонецкаяБахмут", "СумскаяРомны", "ХмельницкаяСтароконстантинов", _
        "ЗапорожскаяДнепрорудное", "ХмельницкаяКамянец-Подольский", "СумскаяШостка", "ДонецкаяПокровск", _
        "ДнепропетровскаяСофиевка", "ЛуганскаяРубежное", "ДонецкаяТорез", "ДонецкаяКраматорск", "ПолтавскаяХорол", _
        "ЗапорожскаяБердянск", "ЛуганскаяКременная")
    For x = LBound(UAName) To UBound(UAName)
    Selection.Replace What:=UAName(x), Replacement:=RuName(x), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    '------формат занятости-----------------------
    Columns(34).Select
    Selection.Replace What:="*резерв*", Replacement:="резерв", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------Сторона -----------------------
    Columns(3).Copy
    Columns(9).PasteSpecial Paste:=xlPasteAll
    Selection.Replace What:="*А*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="*Б*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
        
  
    '------город---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------тип плоскости------------------
Const ColtoFilter2 As Integer = 6
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("Условия").Range("AC2:AC10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------занятость-----------------
Const ColtoFilter4 As Integer = 34
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("Занятость").Range("U2:U4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------начало диапазона-----------------
Set startCell = ws.Range("a1")

'------------убираем автофильтрацию, если таковая присутствует----------
ws.AutoFilterMode = False

'------------определяем диапазон финальной талбицы----------------
Set rngFree = startCell.CurrentRegion

'------------фильтруем и копируем данные-----------
With rngFree

        '------------фильтр по городу----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------фильтр по типу----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------фильтр по занятости----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------копия финального результата----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------создаем новую книгу для внесения финального диапазона----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) 'перенос ширины столбцов - необязательно
'        .Rows(2).Copy
'        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'конец переноса ширины столбцов
    End With
'        '-------------------удалить дубликаты--------------------
'    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
'    For i = lLastRow To 2 Step -1
'        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 7).Value = Cells(i - 1, 7).Value Then
'            Rows(i).Delete
'        End If
'    Next i
'    '-----------------добавляем себестоимость------------------
'    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
'    Cells(lLastRow, 10).Select
'    For i = lLastRow To 2 Step -1
'        If Cells(i, 4).Value = "биллборд" _
'            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("Скидки").Range("AM3") * Cells(i, 11) _
'            Else: If Cells(i, 4).Value = "ситилайт" _
'            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("Скидки").Range("AM4") * Cells(i, 11) _
'            Else Cells(i, 10).Value = ThisWorkbook.Worksheets("Скидки").Range("AM5") * Cells(i, 11)
'    Next



    '-----сохранить выборку------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------снять фильтр в исходном файле----------------

ws.AutoFilterMode = False
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lLastCol = Cells.SpecialCells(xlLastCell).Column
Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
'ActiveWorkbook.Close
    Windows(nameOfGeneralFile).Activate
        Sheets(nameOfSheet2).Select
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
    Windows(nameOfFile).Close
    wb.Close


Set rngFree = Nothing
Set startCell = Nothing
Set ws = Nothing

End Sub
Private Sub HomeAllSheets() 'скроллинг
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = True
    ws.Select
    Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
Next ws
End Sub

Sub MediaPlan()
    Dim iTimer As Single
        iTimer = Timer
    Dim nameOfGeneralFile As String
    Dim nameOfPathGeneralFile As String
    nameOfPathGeneralFile = ActiveWorkbook.Path
    nameOfGeneralFile = ActiveWorkbook.Name
    Application.ScreenUpdating = False 'отключение обновления экрана
    Workbooks.Application.DisplayAlerts = False ' отключение всплывающих окон
    Workbooks(nameOfGeneralFile).Save
    
    Call HomeAllSheets
    
    If ActiveWorkbook.Worksheets("Скидки").Range("E3").Value = "+" Then Call Prime("PrimeNet.xlsx", "Прайм", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E4").Value = "+" Then Call Bigmedia("Bigmedia.xlsx", "Bigmedia", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E5").Value = "+" Then Call Octagon("Octagon.xlsx", "Сверка-Статус Клиенту", "Octagon", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E8").Value = "+" Then Call SVO_news("SVO.xlsx", "SVO", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E6").Value = "+" Then Call Perekhid("Perekhid.xlsx", "Perekhid", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E9").Value = "+" Then Call Luvers("Luvers.xlsx", "Luvers", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E7").Value = "+" Then Call Dovira("Dovira.xlsx", "Dovira_price.xlsx", "Dovira", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E13").Value = "+" Then Call RTM("RTM.xlsx", "RTM", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E18").Value = "+" Then Call Tristar("Tristar.xlsx", "Tristar_GRP.xlsx", "Tristar", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E14").Value = "+" Then Call Sean("Sean_city.xlsx", "Sean_board.xlsx", "Sean", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E11").Value = "+" Then Call Mallis("Mallis.xlsx", "Mallis_GRP.xlsx", "Mallis", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E10").Value = "+" Then Call Alhor("Alhor.xlsx", "Alhor", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E17").Value = "+" Then Call CityDnepr("CityDnepr.xlsx", "CityDnepr", "3.0х6.0", "1.2х1.8", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E19").Value = "+" Then Call Prospect("Prospect.xlsx", "Prospect_GRP.xlsx", "Prospect", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E12").Value = "+" Then Call Megapolis("Megapolis_Kh.xlsx", "Megapolis_Dp.xlsx", "Megapolis_GRP.xlsx", "Megapolis", "3х6", "1.2х1.8 2х3", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E15").Value = "+" Then Call Bomond("Bomond.xlsx", "Bomond_GRP.xlsx", "Bomond", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E20").Value = "+" Then Call ThreeSixDnepr("3x6Dnepr.xlsx", "3x6Dnepr", "Oblast", "Dnipro", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E21").Value = "+" Then Call NashaSprava("NashaSprava_board.xlsx", "NashaSprava", "Львов", "NashaSprava_citylight.xlsx", "NashaSprava_scroll.xlsx", "NashaSprava_GRP.xlsx", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E22").Value = "+" Then Call MegapolisUA("Megapolis_UA.xlsx", "Украина", "Megapolis", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("Скидки").Range("E23").Value = "+" Then Call T52("T52.xlsx", "Лист1", "T52", nameOfPathGeneralFile, nameOfGeneralFile)
    
    Call HomeAllSheets
    Windows(nameOfGeneralFile).Activate ' активация книги откуда запущен макрос
    Sheets("Медиа план").Select
    Dim XCell As Object
    Dim XCol, XRow As Integer
    txtCol = "$$@@2"  ' метка для столбца
    Set XCell = ThisWorkbook.Sheets("Медиа план").Cells.Find(txtCol)
    XCol = XCell.Column
    XRow = XCell.Row
'--------Прайм------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E3").Value = "+" Then
        Range("AL19:AM19").Select
        Selection.AutoFill Destination:=Range("AL19:" & "AM" & XRow - 2), Type:=xlFillDefault
        Range("AL20:" & "AM" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------Бигмедиа------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E4").Value = "+" Then
        Range("AQ19:AR19").Select
        Selection.AutoFill Destination:=Range("AQ19:" & "AR" & XRow - 2), Type:=xlFillDefault
        Range("AQ20:" & "AR" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------РТМ------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E13").Value = "+" Then
        Range("AV19:AW19").Select
        Selection.AutoFill Destination:=Range("AV19:" & "AW" & XRow - 2), Type:=xlFillDefault
        Range("AV20:" & "AW" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------октагон------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E5").Value = "+" Then
        Range("BA19:BB19").Select
        Selection.AutoFill Destination:=Range("BA19:" & "BB" & XRow - 2), Type:=xlFillDefault
        Range("BA20:" & "BB" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------перехид------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E6").Value = "+" Then
        Range("BF19:BG19").Select
        Selection.AutoFill Destination:=Range("BF19:" & "BG" & XRow - 2), Type:=xlFillDefault
        Range("BF20:" & "BG" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------довира------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E7").Value = "+" Then
        Range("BK19:BL19").Select
        Selection.AutoFill Destination:=Range("BK19:" & "BL" & XRow - 2), Type:=xlFillDefault
        Range("BK20:" & "BL" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------ньюсы------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E8").Value = "+" Then
        Range("BP19:BQ19").Select
        Selection.AutoFill Destination:=Range("BP19:" & "BQ" & XRow - 2), Type:=xlFillDefault
        Range("BP20:" & "BQ" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------луверс------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E9").Value = "+" Then
        Range("BV19:BW19").Select
        Selection.AutoFill Destination:=Range("BV19:" & "BW" & XRow - 2), Type:=xlFillDefault
        Range("BV20:" & "BW" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------альхор------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E10").Value = "+" Then
        Range("CA19:CB19").Select
        Selection.AutoFill Destination:=Range("CA19:" & "CB" & XRow - 2), Type:=xlFillDefault
        Range("CA20:" & "CB" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------маллис------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E11").Value = "+" Then
        Range("CF19:CG19").Select
        Selection.AutoFill Destination:=Range("CF19:" & "CG" & XRow - 2), Type:=xlFillDefault
        Range("CF20:" & "CG" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------мегаполис------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E12").Value = "+" Or ActiveWorkbook.Worksheets("Скидки").Range("E22").Value = "+" Then
        Range("CL19:CM19").Select
        Selection.AutoFill Destination:=Range("CL19:" & "CM" & XRow - 2), Type:=xlFillDefault
        Range("CL20:" & "CM" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------сеан------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E14").Value = "+" Then
        Range("CR19:CS19").Select
        Selection.AutoFill Destination:=Range("CR19:" & "CS" & XRow - 2), Type:=xlFillDefault
        Range("CR20:" & "CS" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------бомонд------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E15").Value = "+" Then
        Range("CW19:CX19").Select
        Selection.AutoFill Destination:=Range("CW19:" & "CX" & XRow - 2), Type:=xlFillDefault
        Range("CW20:" & "CX" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------черный квадрат------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E16").Value = "+" Then
        Range("DB19:DC19").Select
        Selection.AutoFill Destination:=Range("DB19:" & "DC" & XRow - 2), Type:=xlFillDefault
        Range("DB20:" & "DC" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------ситиднепр------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E17").Value = "+" Then
        Range("DH19:DI19").Select
        Selection.AutoFill Destination:=Range("DH19:" & "DI" & XRow - 2), Type:=xlFillDefault
        Range("DH20:" & "DI" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------тристар------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E18").Value = "+" Then
        Range("DM19:DN19").Select
        Selection.AutoFill Destination:=Range("DM19:" & "DN" & XRow - 2), Type:=xlFillDefault
        Range("DM20:" & "DN" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------проспект------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E19").Value = "+" Then
        Range("DR19:DS19").Select
        Selection.AutoFill Destination:=Range("DR19:" & "DS" & XRow - 2), Type:=xlFillDefault
        Range("DR20:" & "DS" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------3на6 днепр------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E20").Value = "+" Then
        Range("DW19:DX19").Select
        Selection.AutoFill Destination:=Range("DW19:" & "DX" & XRow - 2), Type:=xlFillDefault
        Range("DW20:" & "DX" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------T52------------
    If ActiveWorkbook.Worksheets("Скидки").Range("E23").Value = "+" Then
        Range("EC19:ED19").Select
        Selection.AutoFill Destination:=Range("EC19:" & "ED" & XRow - 2), Type:=xlFillDefault
        Range("EC20:" & "ED" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If

    
    
    Call HidSheets
    Sheets("Медиа план").Activate
    Range("a1").Select

    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Time finish makros: " & Format((Timer - iTimer) / 86400, "Long Time"), vbExclamation, "" ' таймер час-мин-сек
    
End Sub
Private Sub HidSheets()
Dim wsh As Worksheet, NoHid, i As Long, j As Long
NoHid = Array("Медиа план", "Скидки", "GRP", "План_факт")    'скрыть все листы кроме указанных
For Each wsh In ThisWorkbook.Worksheets
    j = 0
    For i = 0 To UBound(NoHid)
        If wsh.Name <> NoHid(i) Then j = j + 1
    Next i
    If j > UBound(NoHid) Then wsh.Visible = False
Next wsh
End Sub



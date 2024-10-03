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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Sheets(nameOfSheet1).Activate
    ActiveSheet.AutoFilterMode = False
    txtCol1 = "�����"
    txtCol2 = "���"
    txtCol3 = "������"

    Set XCell = Sheets(nameOfSheet1).Cells.Find(txtCol1)
    Set YCell = Sheets(nameOfSheet1).Cells.Find(txtCol2)
    Set ZCell = Sheets(nameOfSheet1).Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
    
    '------������� ���� ����---------
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
            And Cells(i, YCol + 1).Value = "������" Or Cells(i, YCol + 1).Value = "����-����") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "������") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "������" And Cells(i, YCol + 2).Value = "3x6") _
            Then Cells(i, YCol).Value = "����" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------������� ���������� own------

    Columns(16).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 16) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f3").Copy
    Range(Cells(2, 16), Cells(lLastRow, 16)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '----------�������������� � �����--------
    With ActiveSheet.UsedRange
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With
    
    '------�����---------
    
Const ColtoFilter1 As Integer = 4
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 7
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("j2:j10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 9
    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("D1:D40")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 15
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("B2:B4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
'Set StartCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Range(Cells(XRow, 1))
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
        '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 6).Value = Cells(i - 1, 6).Value And Cells(i, 10).Value = Cells(i - 1, 10).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Sheets(nameOfSheet1).Activate
    ActiveSheet.AutoFilterMode = False
    txtCol1 = "�����"
    txtCol2 = "����"

    Set XCell = Sheets(nameOfSheet1).Cells.Find(txtCol1)
    Set YCell = Sheets(nameOfSheet1).Cells.Find(txtCol2)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    
    '--------������ ��� �������----------
    'Cells.MergeCells = False '������ ����������� �����
    'Range("A1:L1").Select
    'Selection.Copy
    'Range("A2").Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    '    :=False, Transpose:=False
    'Rows("1:1").Select
    'Application.CutCopyMode = False
    'Selection.Delete Shift:=xlUp
  
    '------������� ���� ��� ����---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "����������" Then Cells(i, YCol).Value = "��������" Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '---------��������� �������� ������---------
    
    Columns(XCol).Select
    Selection.Replace What:="������", Replacement:="�����", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    '------�����---------
    
Const ColtoFilter1 As Integer = 3
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 8

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("K2:K10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 6

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("E1:E6")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 18

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("C2:C3")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a2")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        '����� �������� ������ ��������
    End With
    
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row

        '-------------------������� ���������--------------------

    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value = Cells(i - 1, 3).Value And Cells(i, 4).Value = Cells(i - 1, 4).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----------------��������� �������������------------------
    txtCol = "Price"  ' ����� ��� �������
    Set ZCell = ActiveSheet.Cells.Find(txtCol)
    If ZCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Price Bigmedia, "
    Else
    ZCol = ZCell.Column
    ZRow = ZCell.Row
    Columns(ZCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, ZCol).Select
    Cells(1, ZCol) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, ZCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = "������" And Cells(i, 3).Value = "����" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB6") * Cells(i, ZCol + 1) _
            Else If Cells(i, 8).Value = "������" And Cells(i, 3).Value = "�������" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB8") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "������" And Cells(i, 3).Value = "������" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB10") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "������" And (Cells(i, 3).Value <> "����" Or Cells(i, 2).Value <> "������" Or Cells(i, 2).Value <> "�������") _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB4") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "��������" And Cells(i, 3).Value = "����" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB7") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "��������" And Cells(i, 3).Value = "�������" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB9") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "��������" And Cells(i, 3).Value = "������" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB11") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "��������" And (Cells(i, 3).Value <> "����" Or Cells(i, 3).Value <> "������" Or Cells(i, 2).Value <> "�������") _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB5") * Cells(i, ZCol + 1) _
            Else: If Cells(i, 8).Value = "��������" _
            Then Cells(i, ZCol).Value = ThisWorkbook.Worksheets("������").Range("AB3") * Cells(i, ZCol + 1)
    Next
    End If

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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

'---------������� ������ ������-----------
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet2).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "�����"
    txtCol2 = "������"

    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol1)
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    XCol = XCell.Column
    YCol = YCell.Column
    
    '------������� ���� ����---------
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "�������� 1.2�1.8[CF]" Or Cells(i, YCol + 1).Value = "�������� 1.2�1.8 [CF]") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "��� 3x6 [BB]" Or Cells(i, YCol + 1).Value = "���������� 3�6 [BB]") _
            Then Cells(i, YCol).Value = "��������" _
            Else If Cells(i, YCol + 1).Value = "�������� 2.3�3.14 [BO]" _
            Then Cells(i, YCol).Value = "������" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
  
    
    '-------��������� ������� ��� ��������������------
    Columns(10).Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    Cells(1, 10) = "�������������"
        
    '----------�������������� � �����--------
    With ActiveSheet.UsedRange.Columns(16)
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With
    
    '------�����---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 4
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("l2:l10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 21
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("d2:d4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
        '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 7).Value = Cells(i - 1, 7).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----------------��������� �������������------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, 10).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = "��������" _
            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("������").Range("AM3") * Cells(i, 11) _
            Else: If Cells(i, 4).Value = "��������" _
            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("������").Range("AM4") * Cells(i, 11) _
            Else Cells(i, 10).Value = ThisWorkbook.Worksheets("������").Range("AM5") * Cells(i, 11)
    Next



    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    txtCol1 = "������"
    txtCol2 = "�����������"
    txtCol3 = "������"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    
    XCol = XCell.Column
    XRow = XCell.Row
    Rows("1:" & XRow - 1).Select
    Selection.Delete Shift:=xlUp
    '------������� ���� ��� ����---------
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
        If (Cells(i, YCol + 1).Value = "����-����" Or Cells(i, YCol + 1).Value = "����-������") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "������") _
            Then Cells(i, YCol).Value = "��������" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'-----------������ ������--------------
    Columns(3).Select
    Selection.Replace What:="�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    '------�����---------
    
Const ColtoFilter1 As Integer = 4
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 13

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("O2:O10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 15

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("i2:i6")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 18

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("G1:G4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        '����� �������� ������ ��������
    End With
        '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 2).Value = Cells(i - 1, 2).Value And Cells(i, 3).Value = Cells(i - 1, 3).Value Then
            Rows(i).Delete
        End If
    Next i
        '-------������� ���������� own------

    Columns(15).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 15) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, 15).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, 11).Value = "������" And Cells(i, 4).Value = "����" And Cells(i, 14).Value = 1 _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("������").Range("AG6") _
            Else If Cells(i, 11).Value = "������" And Cells(i, 4).Value = "����" And Cells(i, 14).Value = 2 _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("������").Range("AG7") _
            Else: If Cells(i, 11).Value = "������" And Cells(i, 4).Value = "����" And Cells(i, 14).Value = 3 _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("������").Range("AG8") _
            Else: If Cells(i, 11).Value = "��������" And Cells(i, 4).Value = "����" _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("������").Range("AH3") * Cells(i, 16) _
            Else: If Cells(i, 11).Value = "��������" And Cells(i, 4).Value = "����" _
            Then Cells(i, 15).Value = ThisWorkbook.Worksheets("������").Range("AH4") * Cells(i, 16) _
            Else: Cells(i, 15).Value = ThisWorkbook.Worksheets("������").Range("AH5") * Cells(i, 16)
    Next


    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False
    
    txtCol3 = "������"
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    ZRow = ZCell.Row
    ZCol = ZCell.Column
    Rows("1:" & ZRow - 1).Select
    Selection.Delete Shift:=xlUp
    
    txtCol1 = "�����"
    txtCol2 = "�����������"


    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
  
    '------������� ���� ��� ����---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol + 1).Value = "����������" Or Cells(i, YCol + 1).Value = "��������" Or _
                (Cells(i, YCol + 1).Value = "������" And Cells(i, YCol + 2).Value = "1.8x1.2")) _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "������") _
            Then Cells(i, YCol).Value = "��������" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------�������������� ������-----------
    Columns(XCol).Select
    Selection.Replace What:="���������� (������������ )", Replacement:="������������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    
    '-------������� ���������� own------

    Columns(11).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 11) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f6").Copy
    Range(Cells(2, 11), Cells(lLastRow, 11)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 4

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("m2:m10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 6

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("g2:g6")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 15

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("e1:e4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a2")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        '����� �������� ������ ��������
    End With
    '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Activate
    ActiveSheet.AutoFilterMode = False
    Rows("1:4").Select
    Selection.Delete Shift:=xlUp
    
    txtCol1 = "�����"
    txtCol2 = "�����������"
    txtCol3 = "������"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
  
    '------������� ���� ��� ����---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "��������" _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "������") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "������" And Cells(i, YCol + 2).Value = "1.2x1.8") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "������" And Cells(i, YCol + 2).Value = "6x3") _
            Then Cells(i, YCol).Value = "����" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------�������������� ������-----------
    Columns(XCol).Select
    Selection.Replace What:="���", Replacement:="����", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    
    '-------������� ���������� own------

    Columns(10).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 10) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f9").Copy
    Range(Cells(2, 10), Cells(lLastRow, 10)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 4

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("p2:p10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 6

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("j2:j10")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 12

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("H1:H4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a2")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        '����� �������� ������ ��������
    End With
        '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value = Cells(i - 1, 3).Value And Cells(i, 7).Value = Cells(i - 1, 7).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------������� �����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2) '�������� ����� Price
'------------���������� ���� ��� ������------------

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
    Cells(1, 14) = "�������������"
    For i = lLastRow To 2 Step -1
        
                Cells(i, 14) = Application.VLookup(Cells(i, 1), Workbooks(nameOfFile2).Sheets("Price").Range( _
                                                                    Workbooks(nameOfFile2).Sheets("Price").Cells(1, 1), _
                                                                    Workbooks(nameOfFile2).Sheets("Price").Cells(lLastRow, 6)), 6, False)
    Next
    Columns(1).Delete
    
    txtCol1 = "�����"
    txtCol2 = "���"
    txtCol3 = "������"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
  
    '------������� ���� ��� ����---------
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "����-����" _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "������") _
            Then Cells(i, YCol).Value = "��������" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------�������������� ������-----------
    Columns(XCol).Select
    Selection.Replace What:="�������", Replacement:="�������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        '----------�������������� ������--------
    Columns(15).Select
    Selection.Replace What:="�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
       
    '------�����---------
    
Const ColtoFilter1 As Integer = 2
    
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A200")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 12

    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("N2:N10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 5

    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("h2:h14")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 16

    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("f1:f4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1)
        '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths
        '����� �������� ������ ��������
    End With
    '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 11).Value = Cells(i - 1, 11).Value And Cells(i, 15).Value = Cells(i - 1, 15).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile

Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False
    txtCol1 = "�����"
    txtCol2 = "���"
    txtCol3 = "������"

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    Set YCell = ActiveSheet.Cells.Find(txtCol2)
    Set ZCell = ActiveSheet.Cells.Find(txtCol3)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    ZCol = ZCell.Column
    '-------������� ������ ������ � �������-----------
    Range(Cells(XRow + 1, 16), Cells(XRow + 1, 39)).Select
    Selection.Copy
    Range(Cells(XRow, 16), Cells(XRow, 39)).Select
    Selection.PasteSpecial xlPasteAll
    Rows(XRow + 1).Select
    Rows(XRow + 1).Delete
    '-------������� ������� � ��������-----------
    Columns(ZCol).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------������� ���� ����---------
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
        If Cells(i, YCol + 1).Value = "����-����" _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "������" And Cells(i, YCol + 2).Value = "1.86x1.3" Or _
            Cells(i, YCol + 2).Value = "1.8x1.2" Or Cells(i, YCol + 2).Value = "1.7x1.2" Or Cells(i, YCol + 2).Value = "1.86x1.27") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "���" Or Cells(i, YCol + 1).Value = "������") _
            And (Cells(i, YCol + 2).Value = "3x6") _
            Then Cells(i, YCol).Value = "��������" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
   
    
    '-------������� ���������� own------

    Columns(16).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 16) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f13").Copy
    Range(Cells(2, 16), Cells(lLastRow, 16)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '----------�������������� ������--------
    Columns(12).Select
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------�����---------
    
Const ColtoFilter1 As Integer = 5
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A175")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 9
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("t2:t10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'---------------������� ����������-------------
Const ColtoFilter3 As Integer = 11
    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("n1:n10")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 18
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("l2:l4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 12).Value = Cells(i - 1, 12).Value Then
            Rows(i).Delete
        End If
    Next i
    
        '----------�������������� � �����--------
    With ActiveSheet.UsedRange.Columns(15)
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)
    ActiveSheet.AutoFilterMode = False
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "�����"
    txtCol2 = "��� ������ "

    Set XCell = ActiveSheet.Cells.Find(txtCol1)
    
    XCol = XCell.Column
    XRow = XCell.Row
    '------������� ���� ����---------
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
        If (Cells(i, YCol + 1).Value = "������ 6�3" _
        Or Cells(i, YCol + 1).Value = "��� 6,2�3,2" _
        Or Cells(i, YCol + 1).Value = "��� 6�3" _
        Or Cells(i, YCol + 1).Value = "��� 5,7�2,5" _
        Or Cells(i, YCol + 1).Value = "��� 5,9� 2,9") _
        Then Cells(i, YCol).Value = "��������" _
        Else: If Cells(i, YCol + 1).Value = "����-���� 1,2x1,8" _
        Then Cells(i, YCol).Value = "��������" _
        Else: If Cells(i, YCol + 1).Value = "������ 3,14�2,32" _
        Then Cells(i, YCol).Value = "������" _
        Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
    '-------������� ���������� own------

    Columns(17).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 17) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f18").Copy
    Range(Cells(2, 17), Cells(lLastRow, 17)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '----------�������������� ������--------
    Columns(15).Select
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
            Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.FormulaR1C1 = "A"
    
 '---------��������� GRP--------------
    Cells(1, 16).EntireColumn.Insert
    Cells(1, 16) = "GRP"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        Cells(i, 16) = Application.IfError(Application.VLookup(Cells(i, 10), Workbooks(nameOfFile2).Sheets("GRP").Range( _
                                                                    Workbooks(nameOfFile2).Sheets("GRP").Cells(1, 10), _
                                                                    Workbooks(nameOfFile2).Sheets("GRP").Cells(lLastRow, 13)), 4, False), "-")
    Next

    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A175")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 3
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("y2:y10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 20
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("q2:q4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 13).Value = Cells(i - 1, 13).Value And Cells(i, 15).Value = Cells(i - 1, 15).Value Then
            Rows(i).Delete
        End If
    Next i
    

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Add
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)
    ActiveSheet.AutoFilterMode = False
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False
'-----------ID Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "ID"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ID Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------��� ����� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "��� Doors"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "��� ����� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("B1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("C1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("D1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------���� 1 Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "���� 1"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���� 1 Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("E1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("F1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("H1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------��� �������� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "��� ��������"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "��� �������� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("J1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------OTS Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "OTS"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "OTS Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("N1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------GRP Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "GRP"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "GRP Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("O1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------���� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("P1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------�������� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "��������"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�������� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("Q1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------���� ����� � ��������� Board-----------
    Workbooks(nameOfFile2).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "���� �����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���� ����� Board, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile2).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow & ":" & ReturnName(XCol + 22) & lLastRow).Copy
    Workbooks("�����1").Activate
    ActiveWorkbook.ActiveSheet.Range("S1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------������ ������--------------
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("L:L").Select
    ActiveSheet.Paste
    Columns("L:L").Select
    Selection.Replace What:="* �", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="* �", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="* �", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'---------������ ����� �����-----------------------

    Workbooks("�����1").Activate
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("I:I").Select
    ActiveSheet.Paste
    Columns("I:I").Select
    Dim fndList As Variant
    Dim x As Long
    fndList = Array("��� ������� 6�3�", "������ ������� 6�3�", "����-������ 6�3�", "������ VIP 6�3�", "������ 6�3�", "���� 6�3�", "��� 6�3�")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:="��������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
'-----------������ �������������---------------
    Workbooks("�����1").Activate
    Columns("J:J").Select
    Selection.Copy
    Columns("R:R").Select
    ActiveSheet.Paste
    Dim Rng As Range
    Dim InputRng As Range, ReplaceRng As Range
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set InputRng = ActiveWorkbook.ActiveSheet.Range("R1:" & "R" & lLastRow)
    Set ReplaceRng = ThisWorkbook.Sheets("������").Range("I2:j8")
    For Each Rng In ReplaceRng.Columns(1).Cells
        InputRng.Replace What:=Rng.Value, Replacement:=Rng.Offset(0, 1).Value
    Next
    
 '-----------�������� �������� �������� �� ����� City � TYPE-----------
    Workbooks("�����1").Activate
    Range("G1").Value = "����"
    Range("I1").Value = "TYPE"
    Range("K1").Value = "������"
    Range("L1").Value = "�������"
    Range("M1").Value = "��"
    Range("R1").Value = "�������������"
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
    txtCol = "ID"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "ID City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------��� Doors City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "��� Doors"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "��� Doors City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("B1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("C1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("D1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------���� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("G1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("H1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------��� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "���"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "��� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("J1").PasteSpecial Paste:=xlPasteAll
    End If
 '-----------������ City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "������"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "������ City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("K1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------������� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�������"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "������� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("L1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------�� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "��"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("M1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------OTS City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "OTS"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "OTS City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("N1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------GRP City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "GRP"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "GRP City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("O1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------���� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("P1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------�������� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "��������"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�������� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("Q1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------����� � ��������� City-----------
    Workbooks(nameOfFile).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� City, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol) & XRow + 1 & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks("�����2").Activate
    ActiveWorkbook.ActiveSheet.Range("S1").PasteSpecial Paste:=xlPasteAll
'-----------���������---------------
    Workbooks(nameOfFile).Activate
    ActiveWorkbook.ActiveSheet.Range(ReturnName(XCol + 2) & XRow + 1 & ":" & ReturnName(XCol + 23) & lLastRow).Copy
    Workbooks("�����2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range("T1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------��������� ������--------------
    Columns("L:L").Select
    Selection.Replace What:="A*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="B*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'------�������������� ������ � ���� � ��������-----------
    Workbooks("�����2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 1 Step -1
        If Cells(i, 10).Value = "����-����" _
        Or (Cells(i, 10).Value = "������" And Cells(i, 11).Value = "1.8x1.2") _
        Then Cells(i, 9).Value = "��������" _
        Else If (Cells(i, 10).Value = "������" And Cells(i, 11).Value = "3x1.5") _
        Then Cells(i, 9).Value = "���" _
        Else: If (Cells(i, 10).Value = "������" And Cells(i, 11).Value = "3x6") _
        Then Cells(i, 9).Value = "��������" _
        Else: Cells(i, 9).Value = Cells(i, 10)
    Next
'------�������� �������������-----------
    Workbooks("�����2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 1 Step -1
        If (Cells(i, 9).Value = "��������" And Cells(i, 8).Value = "*�����������*") _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("������").Range("J10") _
        Else: If (Cells(i, 9).Value = "��������" And Cells(i, 8).Value = "*�������*") _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("������").Range("J11") _
        Else: If Cells(i, 9).Value = "������" _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("������").Range("J13") _
        Else: If Cells(i, 9).Value = "��������" _
        Then Cells(i, 18).Value = ThisWorkbook.Sheets("������").Range("J14") _
        Else: Cells(i, 18).Value = ThisWorkbook.Sheets("������").Range("J12")
    Next
'-----------��������� �����-------------
    Workbooks("�����2").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Workbooks("�����1").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
'----------�������������� � �����--------
    With ActiveSheet.UsedRange.Columns(15)
        .Replace ",", "."
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With

    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 9
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("U2:U10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 21
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("M2:M4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues
        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
            '-------------------������� ���������--------------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 12).Value = Cells(i - 1, 12).Value Then
            Rows(i).Delete
        End If
    Next i
    '-------------����������� �����------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(2, 3).Select
    Cells(2, 3).Value = "������"
    Cells(2, 3).Select
    Selection.AutoFill Destination:=Range(Cells(2, 3), Cells(lLastRow, 3)), Type:=xlFillDefault

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    Windows("�����1").Close
    Windows("�����2").Close
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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  '�������� �����
    ActiveSheet.AutoFilterMode = False

    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False
    txtCol2 = "��������"

    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------������� ���� ����---------
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
        If Cells(i, YCol + 1).Value = "�����" _
            Then Cells(i, YCol).Value = "������" _
            Else: If (Cells(i, YCol + 1).Value = "������" Or Cells(i, YCol + 1).Value = "���") _
            Then Cells(i, YCol).Value = "��������" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'------�������� �������������-----------
    Columns(YCol + 1).Select
    Selection.Insert Shift:=xlToRight
    Cells(1, YCol + 1) = "�������������"
    For i = lLastRow To 1 Step -1
        If Cells(i, YCol + 2).Value = "���" _
        Then Cells(i, YCol + 1).Value = ThisWorkbook.Sheets("������").Range("M2") _
        Else: If Cells(i, YCol + 2).Value = "������" _
        Then Cells(i, YCol + 1).Value = ThisWorkbook.Sheets("������").Range("M3") _
        Else: If Cells(i, YCol + 2).Value = "�����" _
        Then Cells(i, YCol + 1).Value = ThisWorkbook.Sheets("������").Range("M4")
    Next
    '---------������� ������ ������----------
    Dim r As Long
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For r = LastRow To 2 Step -1
    If Application.CountA(Rows(r)) = 0 Then
        Rows(r).Delete
    End If
    Next r
    '-------��������� �����------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Columns(3).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 3) = "�����"
    Cells(2, 3) = "����"
    Cells(2, 3).Copy
    Range(Cells(3, 3), Cells(lLastRow, 3)).Select
    Selection.PasteSpecial Paste:=xlAll
'-----------������ ������--------------
    Columns(5).Select
    Selection.Replace What:="����-��", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'----------- ������� ������� �� �������--------------
    Columns(4).Select
    Selection.Replace What:="(�5)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�4)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�3)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�2)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�1)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�5)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�4)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�3)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�2)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�1)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(�)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------��������� GRP--------------
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

    '------�����---------
    
Const ColtoFilter1 As Integer = 3
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 6
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("r2:r10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 15
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("j2:j4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
'Set StartCell = Workbooks(nameOfFile).Worksheets(nameOfSheet1).Range(Cells(XRow, 1))
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 5).Value = Cells(i - 1, 5).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Columns("A:BB").Hidden = False '����������� ��� �������
    ActiveWindow.FreezePanes = False '������ ����������� ��������
    txtCol2 = "��� �����������"

    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------������� ���� ����---------
    Rows(1 & ":" & YRow - 1).Select
    Selection.Delete Shift:=xlUp
    Columns(YCol).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    lLastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
    Cells(lLastRow, YCol).Select
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "������ 2,30�3,140" _
            Then Cells(i, YCol).Value = "������" _
            Else: If (Cells(i, YCol + 1).Value = "������ 3�6" Or Cells(i, YCol + 1).Value = "��� 3�6,2" Or _
                Cells(i, YCol + 1).Value = "��� 3�6" Or Cells(i, YCol + 1).Value = "��� 3,2�6,2") _
            Then Cells(i, YCol).Value = "��������" _
            Else: If (Cells(i, YCol + 1).Value = "����-�������� 1.2�1.8" Or Cells(i, YCol + 1).Value = "����-���� 1,2�1,8") _
            Then Cells(i, YCol).Value = "��������" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'------�������� �������������-----------
    Columns(11).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 11) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f10").Copy
    Range(Cells(2, 11), Cells(lLastRow, 11)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
        
    '-------��������� �����------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    Columns(2).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 2) = "�����"
    Cells(2, 2) = "����"
    Cells(2, 2).Copy
    Range(Cells(3, 2), Cells(lLastRow, 2)).Select
    Selection.PasteSpecial Paste:=xlAll
'-----------������ ������--------------
    Columns(4).Select
    Selection.Replace What:="�*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------�����---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 5
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("Q2:Q10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

    '--------------������------------------
Const ColtoFilter3 As Integer = 6
    Set rngSize = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("K2:K10")
    arr3 = Application.WorksheetFunction.Transpose(rngSize.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 16
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("j2:j4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("A1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� �������----------------
        .AutoFilter Field:=ColtoFilter3, Criteria1:=arr3, Operator:=xlFilterValues

        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value = Cells(i - 1, 3).Value And Cells(i, 4).Value = Cells(i - 1, 4).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Sheets.Add
    Workbooks(nameOfFile).Sheets.Add
    
'-----------��������� ��������� �� ����� �����-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetBoard).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "������������������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-------------��������� ��� ���������------------
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(lLastCol).Copy
    Columns(ReturnName(lLastCol + 1) & ":" & ReturnName(lLastCol + 2)).PasteSpecial Paste:=xlPasteAll
    Cells(1, lLastCol + 1).Value = "��� ���������"
    Cells(2, lLastCol + 1).Value = "��������"
    Cells(1, lLastCol + 2).Value = "������"
    Cells(2, lLastCol + 2).Value = "6�3"
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & 2).Copy
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & lLastRow).PasteSpecial Paste:=xlPasteValues
'-----------��������� ��������� �� ����� ���� � �� �����-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetBoard).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�����������������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(XCol) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(lLastCol + 1) & 1).PasteSpecial Paste:=xlPasteAll
    End If
'-----------������� ������ ������� �� ������ ������---------
    For i = 30 To 1 Step -1
        If Cells(1, i).Value = 0 Then
            Columns(i).Delete
            i = i - 1
        End If
    Next i

'-----------��������� ��������� �� ���������� �����-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�����������������������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(XCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����2").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-------------��������� ��� ���������------------
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns(lLastCol).Copy
    Columns(ReturnName(lLastCol + 1) & ":" & ReturnName(lLastCol + 2)).PasteSpecial Paste:=xlPasteAll
    Cells(1, lLastCol + 1).Value = "��� ���������"
    Cells(2, lLastCol + 1).Value = "��������"
    Cells(1, lLastCol + 2).Value = "������"
    Cells(2, lLastCol + 2).Value = "1.2x1.8"
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & 2).Copy
    Range(ReturnName(lLastCol + 1) & 2 & ":" & ReturnName(lLastCol + 2) & lLastRow).PasteSpecial Paste:=xlPasteValues

'-----------��������� ��������� �� ���������� ���� � �� �����-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���������������������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(XCol) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����2").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(lLastCol + 1) & 1).PasteSpecial Paste:=xlPasteAll
    End If
'-----------������� ������ ������� �� ������ ������---------
    For i = 30 To 1 Step -1
        If Cells(1, i).Value = 0 Then
            Columns(i).Delete
            i = i - 1
        End If
    Next i
'-----------��������� �����-------------
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Copy
    Workbooks(nameOfFile).Sheets("����2").Activate
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    '-------������� ���������� own------

    Columns(9).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 9) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f17").Copy
    Range(Cells(2, 9), Cells(lLastRow, 9)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    
'----------������ �������--------
    With ActiveSheet.UsedRange.Columns(13)
        .Replace "�", ""
    End With

'----------�������������� � �����--------
    With ActiveSheet.UsedRange.Columns(13)
        arr = .Value
        .NumberFormat = "General"
        .Value = arr
    End With
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 4
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("x2:x10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 15

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=1
        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    '-------------����������� �����------------
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(2, 1).Select
    Cells(2, 1).Value = "�����"
    Cells(2, 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(lLastRow, 1)), Type:=xlFillDefault

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  '�������� �����
    ActiveSheet.AutoFilterMode = False

    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False '������ ����������� ��������
    Cells.MergeCells = False '������ ����������� �����

    txtCol2 = "���"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------������� ���� ����---------
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
        If Cells(i, YCol + 1).Value = "������ 3,14�2,32" _
            Then Cells(i, YCol).Value = "������" _
            Else: If (Cells(i, YCol + 1).Value = "��� 3�6" Or Cells(i, YCol + 1).Value = "������3�6" Or Cells(i, YCol + 1).Value = "��� 3,2�6,2") _
            Then Cells(i, YCol).Value = "��������" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 1)
    Next
'------�������� �������������-----------
    Columns(12).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 12) = "�������������"
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f19").Copy
    Range(Cells(4, 12), Cells(lLastRow, 12)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
'-----------������ ������--------------
    Columns(7).Select
    Selection.Replace What:="�/�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------��������� GRP--------------
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
    '-------������� ������ ������-----------
    For i = 5 To 1 Step -1
        If Cells(i, 1).Value = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 4
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("z2:z10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 15
Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

                                                           
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=1

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    '-------������� ��������� ��������------------
    Columns(6).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Selection.Replace What:="-*", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 6).Value = Cells(i - 1, 6).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i
    Columns(6).Delete
    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Sheets.Add
    Workbooks(nameOfFile).Sheets.Add
    
'-----------��������� ��������� �� �����-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetBoard).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row

    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�����_���_����_�������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    txtCol = ""
    Set ZCell = Workbooks(nameOfFile).ActiveSheet.Range(ReturnName(1) & XRow & ":" & ReturnName(1) & lLastRow).Find(txtCol)
    ZCol = ZCell.Column
    ZRow = ZCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(lLastCol) & ZRow - 1).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
'-----------��������� ��������� �� ����������-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�����_����_����_�������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
'-----------��������� ������� �����(��� ����������� � �������)-------
    Columns(9).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, 9) = "�����"

'-----------��������� ��������� �� ������-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  '�������� �����
    Workbooks(nameOfFile1).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "�����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�����_���_����_�����, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
    '------������� ���� ����---------
    txtCol = "�a����"  ' ����� ��� �������
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "������_����_�������, "
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
        If (Cells(i, YCol + 1).Value = "3�6" Or Cells(i, YCol + 1).Value = "2,9�5,9") _
            And (Cells(i, YCol + 2).Value = "������" _
            Or Cells(i, YCol + 2).Value = "����" _
            Or Cells(i, YCol + 2).Value = "��������" _
            Or Cells(i, YCol + 2).Value = "��������-������" _
            Or Cells(i, YCol + 2).Value = "����") _
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 1).Value = "3,1�2,2" Or Cells(i, YCol + 1).Value = "3,33�2,3") _
            And (Cells(i, YCol + 2).Value = "������" Or Cells(i, YCol + 2).Value = "�������") _
            Then Cells(i, YCol).Value = "������" _
            Else Cells(i, YCol).Value = Cells(i, YCol + 2)
    Next
    End If
'------�������� �������������-----------
    Columns(12).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 12) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        If (Cells(i, YCol - 1).Value = "*VIP*") _
        Then Cells(i, 12).Value = Cells(i, 13) * ThisWorkbook.Sheets("������").Range("q3") _
        Else: Cells(i, 12).Value = Cells(i, 13) * ThisWorkbook.Sheets("������").Range("q2")
    Next
'-----------������ ������--------------
    Columns(8).Select
    Selection.Replace What:="�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------��������� GRP--------------
     Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)  '�������� �����
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
'------������ �� ���������------------------
    colorReserv = RGB(204, 153, 255)
    colorFree = RGB(255, 255, 255)
    For i = lLastRow To 2 Step -1
            If Cells(i, 15).Interior.Color = colorReserv _
            Then Cells(i, 15).Value = "������" _
            Else If Cells(i, 15).Interior.Color = colorFree _
            Then Cells(i, 15).Value = "��������"
    Next i
    
    Columns(1).Select
    Selection.Replace What:="�a�����", Replacement:="�������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 5
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("s2:s10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 15
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("k2:k4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)


Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues


        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False
    ActiveWindow.FreezePanes = False '������ ����������� ��������
    Cells.MergeCells = False '������ ����������� �����

    txtCol = "������� ���������:"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    
    YCol = YCell.Column
    YRow = YCell.Row
    
    '------������� ���� ����---------
    Columns(YCol).Select
    Application.CutCopyMode = False
    Selection.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Columns(YCol).Select
    '-----------������ �����---------------
    Dim fndList As Variant
    Dim x As Long
    fndList = Array("��� ������", "��� �������", "�����", "���������", "������ VIP", "������")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:="��������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    fndList = Array("������ VIP", "������", "")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:="������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
'-----------������ �������������---------------
    Columns(YCol + 1).Select
    Selection.Copy
    Selection.Insert Shift:=xlToLeft

    Dim Rng As Range
    Dim InputRng As Range, ReplaceRng As Range
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set InputRng = ActiveWorkbook.ActiveSheet.Range(Cells(1, YCol + 2), Cells(lLastRow, YCol + 2))
    Set ReplaceRng = ThisWorkbook.Sheets("������").Range("S3:T14")
    For Each Rng In ReplaceRng.Columns(1).Cells
        InputRng.Replace What:=Rng.Value, Replacement:=Rng.Offset(0, 1).Value
    Next
    Cells(1, YCol + 2).Select
    Cells(1, YCol + 2) = "�������������"
'-----------������ ������--------------
    txtCol = "�������:"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    
    YCol = YCell.Column
    YRow = YCell.Row

    Columns(YCol).Select
    Selection.Replace What:="�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 '---------��������� GRP--------------
    '������������� ���� ���� � �����
    
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ThisWorkbook.Sheets("������").Range("B15").Copy
    Range("B2:" & "B" & lLastRow).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    '��������� �������� GRP
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile1)  '�������� �����
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
    
    '-------------����������� �����------------
    Columns(2).Select
    Selection.Insert Shift:=xlRight
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(2, 2).Select
    Cells(2, 2).Value = "������"
    Cells(2, 2).Select
    Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lLastRow, 2)), Type:=xlFillDefault
    '-------------���������� ���������------------
    txtCol = "��:"
    
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Rows(1).Cells.Find(txtCol)
    
    YCol = YCell.Column
    YRow = YCell.Row

    Columns(YCol).Select
    Selection.Insert Shift:=xlToLeft
    Workbooks(nameOfGeneralFile).Worksheets("������").Range("g15").Copy
    Cells(1, YCol).Select
    Selection.PasteSpecial Paste:=xlPasteAll
    For i = lLastRow To 2 Step -1
        If Cells(i, YCol + 1).Value = "" Then Cells(i, YCol + 1).Value = Cells(1, YCol).Value + 365
            If Cells(i, YCol - 1).Value <= Cells(1, YCol).Value And Cells(1, YCol).Value <= Cells(i, YCol + 1).Value _
                Then Cells(i, YCol) = "��������" _
                Else Cells(i, YCol) = "������"
        
    Next
    '------�����---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 12
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("v2:v10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 6
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("N2:N4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
            Rows(i).Delete
        End If
    Next i
    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Sheets.Add
    
'-----------��������� ��������� �� �������-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetRegion).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "����"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�������_3x6Dnepr, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
    Columns(XCol + 1).Select
    Selection.Insert Shift:=xlToRight
    Cells(1, XCol + 1).Value = "�����"

'-----------��������� ��������� �� ������-----------
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    txtCol = "̳���"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�����_3x6Dnepr, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
    '--------������ ��������---------
    txtCol = "�����"  ' ����� ��� �������
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "������_�������3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Columns(YCol).Select
    Selection.Replace What:="�", Replacement:="x", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
    '------������� ���� ����---------
    txtCol = "������"  ' ����� ��� �������
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "������_�������3x6Dnepr, "
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
            Then Cells(i, YCol).Value = "��������" _
            Else If (Cells(i, YCol + 2).Value = "1,8x1,2" And Cells(i, YCol + 1).Value = "sity-light") _
            Then Cells(i, YCol).Value = "��������" _
            Else: Cells(i, YCol).Value = Cells(i, YCol + 2)
    Next
    End If
'------�������� �������������-----------
    txtCol = "��."  ' ����� ��� �������
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����_�������3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Columns(YCol + 1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, YCol + 1) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Workbooks(nameOfGeneralFile).Worksheets("�������").Range("f20").Copy
    Range(Cells(2, YCol + 1), Cells(lLastRow, YCol + 1)).Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    End If
'-----------������ ������--------------
    txtCol = "��."  ' ����� ��� �������
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�������_�������3x6Dnepr, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Range(Cells(2, YCol), Cells(lLastRow, YCol)).Select
    Selection.Replace What:="�*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
'------������ �� ���������------------------
   
    Columns(1).Select
    Dim fndList, fndCity As Variant
    Dim x As Long
    fndList = Array("�����", "��������� ", "�������������", "���'������", "������������", "ͳ������", "��������")
    fndCity = Array("�����", "���������", "������������", "���������", "������������", "��������", "��������")
    For x = LBound(fndList) To UBound(fndList)
    Selection.Replace What:=fndList(x), Replacement:=fndCity(x), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 3
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("AA2:AA10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 17
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("s2:s4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)


Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues


        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 12).Value = Cells(i - 1, 12).Value And Cells(i, 14).Value = Cells(i - 1, 14).Value Then
            Rows(i).Delete
        End If
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Sheets.Add
    
'-----------��������� ��������� �� �����-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    Workbooks(nameOfFile).Sheets.Add
    Workbooks(nameOfFile).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row

    txtCol = "link"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Link_���_����������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Columns(6).Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlLeft
    Cells(XRow, 6).Value = "Type"
    For i = lLastRow To XRow + 1 Step -1
        If InStr(1, Cells(i, 4), "������") <> 0 _
            Then Cells(i, 6).Value = "������" _
            Else: Cells(i, 6).Value = "��������"
    Next
    Range(ReturnName(1) & XRow & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    ActiveWorkbook.ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    End If
    '-----------��������� ��������� �� ����������-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile2)  '�������� �����
    Workbooks(nameOfFile2).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
            '---������ ������� ������� ��� � ��� ����
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 4).End(xlUp).Row
    txtCol = "link"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile2).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Link_��������_����������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Columns(6).Select
    Selection.Insert Shift:=xlRight
    Cells(XRow, 6).Value = "Type"
    For i = lLastRow To XRow + 1 Step -1
        If InStr(1, Cells(i, 4), "������") <> 0 _
            Then Cells(i, 6).Value = "����������" _
            Else: Cells(i, 6).Value = "��������"
    Next
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If
    '-----------��������� ��������� �� ��������-----------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile3)  '�������� �����
    Workbooks(nameOfFile3).Sheets(nameOfSheetCity).Activate
    ActiveSheet.AutoFilterMode = False
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    txtCol = "link"  ' ����� ��� �������
    Set XCell = Workbooks(nameOfFile3).ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Link_������_����������, "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    Columns(6).Select
    Selection.Insert Shift:=xlRight
    Cells(XRow, 6).Value = "Type"
    Cells(XRow + 1, 6).Value = "������"
    Cells(XRow + 1, 6).Select
    Selection.AutoFill Destination:=Range(Cells(XRow + 1, 6), Cells(lLastRow, 6)), Type:=xlFillDefault
    Range(ReturnName(1) & XRow + 1 & ":" & ReturnName(lLastCol) & lLastRow).Copy
    Workbooks(nameOfFile).Sheets("����1").Activate
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    ActiveWorkbook.ActiveSheet.Range(ReturnName(1) & lLastRow + 1).PasteSpecial Paste:=xlPasteAll
    End If

'------�������� �������������-----------
    Columns(6).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Dim Rng As Range
    Dim InputRng As Range, ReplaceRng As Range
    lLastRow = Cells(Rows.Count, 3).End(xlUp).Row
    Set InputRng = ActiveWorkbook.ActiveSheet.Range(Cells(1, 6), Cells(lLastRow, 6))
    Set ReplaceRng = ThisWorkbook.Sheets("������").Range("V3:W7")
    For Each Rng In ReplaceRng.Columns(1).Cells
        InputRng.Replace What:=Rng.Value, Replacement:=Rng.Offset(0, 1).Value
    Next
    Cells(1, 6).Select
    Cells(1, 6) = "�������������"
    Range(ReturnName(6) & 2 & ":" & ReturnName(6) & lLastRow).Select
    Selection.NumberFormat = "0.00"
'-----------������ �����--------------
    Columns(7).Select
    Selection.Replace What:="������", Replacement:="��������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="����������", Replacement:="��������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'-----------������ ������--------------
    txtCol = "�������"  ' ����� ��� �������
    Set YCell = Rows(1).Cells.Find(txtCol)
    If YCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "�������_����������, "
    Else
    YCol = YCell.Column
    YRow = XCell.Row
    Range(Cells(2, YCol), Cells(lLastRow, YCol)).Select
    Selection.Replace What:="����.", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End If
'-------------����������� �����------------
    Columns(2).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(2, 2).Select
    Cells(2, 2).Value = "�����"
    Cells(2, 2).Select
    Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lLastRow, 2)), Type:=xlFillDefault
 '---------��������� GRP--------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile4)  '�������� �����
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
    
    '------�����---------
    
Const ColtoFilter1 As Integer = 2
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 9
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("AB2:AB10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 11
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("T2:T4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)
'--------��������� ���� � �������-----------------
Const ColtoFilter5 As Integer = 5
    
Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:="<>*������*", Operator:=xlFilterValues
        '------------������ ���������� ��� � ��������----------------
        .AutoFilter Field:=ColtoFilter5, Criteria1:="<>*�.*", Operator:=xlAnd, Criteria2:="<>*�.*"


        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        '.Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        ws2.Cells(2, 1).PasteSpecial Paste:=xlPasteAll 'xlPasteColumnWidths'����� �������� ������ ��������
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
    End With
    
    '-------------------������� ���������--------------------
    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
    For i = lLastRow To 2 Step -1
        If Cells(i, 3).Value <> 0 Then _
            If Cells(i, 3).Value = Cells(i - 1, 3).Value _
            Then Rows(i).Delete
            
    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
    
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "�����"
    txtCol2 = "���"

    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol1)
    Set YCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol2)
    
    XCol = XCell.Column
    XRow = XCell.Row
    YCol = YCell.Column
    
    '------������ �������---------
    Rows("1:" & (XRow - 1)).Select
    Selection.Delete Shift:=xlUp
    
    Cells(1, YCol).Select
    Cells(1, YCol) = "Type"
    Columns(YCol).Select
    Selection.Replace What:="��������", Replacement:="��������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '-----------������ ������--------------
    Columns(YCol + 1).Select
    Selection.Replace What:="�", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="�", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'------�������� �������������-----------
    Columns(12).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 12) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        Cells(i, 12).Value = Cells(i, 11) * (1 - ThisWorkbook.Sheets("������").Range("B22"))
    Next

    '------������ �� ���������------------------
    colorReserv = RGB(204, 153, 255)
    colorFree = RGB(255, 255, 255)
    For i = lLastRow To 2 Step -1
            If Cells(i, 13).Interior.Color = colorReserv _
            Then Cells(i, 13).Value = "������" _
            Else If Cells(i, 13).Interior.Color = colorFree _
            Then Cells(i, 13).Value = "��������"
    Next i
    '------��������� ���� �� ���------------------
    Columns(1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    For i = lLastRow To 2 Step -1
            Cells(i, 1).Value = Cells(i, 2).Value & Cells(i, 4).Value
    Next i
    
    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 8
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("s2:s10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 14
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("k2:k4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)


Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range(Cells(1, 1), Cells(lLastRow, 35))

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues


        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With
        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
        .Rows(2).Copy
        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
    
    '-------------------������� ���������--------------------
'    LastRow = ActiveSheet.UsedRange.Rows.Count - 1 + ActiveSheet.UsedRange.Row
'    For i = lLastRow To 2 Step -1
'        If Cells(i, 4).Value = Cells(i - 1, 4).Value And Cells(i, 8).Value = Cells(i - 1, 8).Value Then
'            Rows(i).Delete
'        End If
'    Next i

    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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

'---------������� ������ ������-----------
Windows(nameOfGeneralFile).Activate
    Sheets(nameOfSheet2).Select
    lLastCol = Cells.SpecialCells(xlLastCell).Column
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range(Cells(1, 1), Cells(lLastRow, lLastCol)).Clear
    
'--------------�����------------------
    Workbooks.Open (pathDir & "\Setka\" & nameOfFile)  '�������� �����
    ActiveSheet.AutoFilterMode = False

    txtCol1 = "�������"
    
    Set XCell = Workbooks(nameOfFile).ActiveSheet.Cells.Find(txtCol1)
    
    XCol = XCell.Column
    XRow = XCell.Row
    
    '------������---------
    Rows("1:" & (XRow - 1)).Select
    Selection.Delete Shift:=xlUp
   
    Columns(XCol + 2).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, XCol + 2).Select
    Cells(1, XCol + 2) = "Type"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(lLastRow, XCol + 2).Select
    For i = lLastRow To 2 Step -1
        Cells(i, XCol + 2).Value = "��������"
    Next
    
    '------�������� �������������-----------
    Columns(9).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Cells(1, 10) = "�������������"
    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lLastRow To 2 Step -1
        Cells(i, 10).Value = Cells(i, 9) * (1 - ThisWorkbook.Sheets("������").Range("B23"))
    Next
    '------��������� ���� �� ���------------------
    Columns(1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    For i = lLastRow To 2 Step -1
            Cells(i, 1).Value = Cells(i, 4).Value & Cells(i, 5).Value
    Next i
        '--------������� � ���. �� ���.------------
    Columns(1).Select
    Dim RuName, UAName As Variant
    Dim x As Long
    UAName = Array("�������������", "�������������", "����������������-������������� ", "���������������������", "ʳ������������ �����������", _
        "ʳ������������ ���������", "ʳ������������ ����������", "ʳ������������ ����'����", "���������� ����� ����� ", "���������� ��������", "���������� ��������", _
        "��������������� ������������", "��������������� ���������", "��������������� ³���������", "���������� �����", "�����������������", "ʳ������������ ������������", _
        "��������������� ����������", "��������������� �'��������", "��������������� ���� ���� ", "���������� ���������", "������������'�����", "��������������������", _
        "�����������������", "����������������", "��������������", "���������������", "��������������� ͳ������", "����������������", "������� ������", "������� ������", _
        "���������� �������", "�����������������", "������� ��������", "�������������� ", "������������", "��������������������������", _
        "�������������������", "��������������'�����-����������", "�������������", "�����������������", _
        "��������������� ��������", "���������������", "���������������", "�������������������", "���������� �����", _
        "������������������", "����������������")
    RuName = Array("���������������", "���������������", "�����������������-�������������", "����������������������", "�������������������������", _
        "����������������������", "�������������������������", "����������������������", "����������������� ������", "������������������", "������������������", _
        "����������������������������", "�������������������������", "���������������������������", "���������������", "�����������������", "���������������������������", _
        "��������������������������", "�������������������������", "���������������������� ����", "�������������������", "����������������", "����������������������", _
        "�����������������", "������������������", "���������������", "�����������������", "������������������������", "�����������������", "��������������", "��������������", _
        "������������������", "���������������������", "�����������������", "��������������", "������������", "����������������������������", _
        "�����������������������", "������������������-����������", "�������������", "����������������", _
        "������������������������", "�����������������", "�������������", "������������������", "���������������", _
        "�������������������", "������������������")
    For x = LBound(UAName) To UBound(UAName)
    Selection.Replace What:=UAName(x), Replacement:=RuName(x), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next x
    '------������ ���������-----------------------
    Columns(34).Select
    Selection.Replace What:="*������*", Replacement:="������", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '------������� -----------------------
    Columns(3).Copy
    Columns(9).PasteSpecial Paste:=xlPasteAll
    Selection.Replace What:="*�*", Replacement:="A", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="*�*", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
        
  
    '------�����---------
    
Const ColtoFilter1 As Integer = 1
    Set rngCity = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("A2:A350")
    arr1 = Application.WorksheetFunction.Transpose(rngCity.Value)
    
    '--------------��� ���������------------------
Const ColtoFilter2 As Integer = 6
    Set rngType = Workbooks(nameOfGeneralFile).Worksheets("�������").Range("AC2:AC10")
    arr2 = Application.WorksheetFunction.Transpose(rngType.Value)

'--------���������-----------------
Const ColtoFilter4 As Integer = 34
    Set rngReserv = Workbooks(nameOfGeneralFile).Worksheets("���������").Range("U2:U4")
    arr4 = Application.WorksheetFunction.Transpose(rngReserv.Value)

Set ws = ActiveSheet

'------------������ ���������-----------------
Set startCell = ws.Range("a1")

'------------������� ��������������, ���� ������� ������������----------
ws.AutoFilterMode = False

'------------���������� �������� ��������� �������----------------
Set rngFree = startCell.CurrentRegion

'------------��������� � �������� ������-----------
With rngFree

        '------------������ �� ������----------------
        .AutoFilter Field:=ColtoFilter1, Criteria1:=arr1, Operator:=xlFilterValues
                                                            
        '------------������ �� ����----------------
        .AutoFilter Field:=ColtoFilter2, Criteria1:=arr2, Operator:=xlFilterValues
        
        '------------������ �� ���������----------------
        .AutoFilter Field:=ColtoFilter4, Criteria1:=arr4, Operator:=xlFilterValues

        '------------����� ���������� ����������----------------
        .Offset(1, 0).EntireRow.Copy
    
End With

        '------------������� ����� ����� ��� �������� ���������� ���������----------------

Set ws2 = Workbooks.Add(xlWBATWorksheet).Sheets(1)
    With ws.UsedRange
        .Copy ws2.Cells(1, 1) '������� ������ �������� - �������������
'        .Rows(2).Copy
'        ws2.Cells(2, 1).PasteSpecial 8 'xlPasteColumnWidths'����� �������� ������ ��������
    End With
'        '-------------------������� ���������--------------------
'    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
'    For i = lLastRow To 2 Step -1
'        If Cells(i, 8).Value = Cells(i - 1, 8).Value And Cells(i, 7).Value = Cells(i - 1, 7).Value Then
'            Rows(i).Delete
'        End If
'    Next i
'    '-----------------��������� �������������------------------
'    lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
'    Cells(lLastRow, 10).Select
'    For i = lLastRow To 2 Step -1
'        If Cells(i, 4).Value = "��������" _
'            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("������").Range("AM3") * Cells(i, 11) _
'            Else: If Cells(i, 4).Value = "��������" _
'            Then Cells(i, 10).Value = ThisWorkbook.Worksheets("������").Range("AM4") * Cells(i, 11) _
'            Else Cells(i, 10).Value = ThisWorkbook.Worksheets("������").Range("AM5") * Cells(i, 11)
'    Next



    '-----��������� �������------
    Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
    Dim sSuff1$: sSuff1 = Format(Now, "dd.mm")
    ActiveWorkbook.SaveAs Filename:= _
        pathDir & "\Vyborka\" & "Vyborka_" & sSuff1 & "_" & sSuff & "_" & nameOfFile
Set wb = ActiveWorkbook
        
    '------------����� ������ � �������� �����----------------

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
Private Sub HomeAllSheets() '���������
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
    Application.ScreenUpdating = False '���������� ���������� ������
    Workbooks.Application.DisplayAlerts = False ' ���������� ����������� ����
    Workbooks(nameOfGeneralFile).Save
    
    Call HomeAllSheets
    
    If ActiveWorkbook.Worksheets("������").Range("E3").Value = "+" Then Call Prime("PrimeNet.xlsx", "�����", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E4").Value = "+" Then Call Bigmedia("Bigmedia.xlsx", "Bigmedia", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E5").Value = "+" Then Call Octagon("Octagon.xlsx", "������-������ �������", "Octagon", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E8").Value = "+" Then Call SVO_news("SVO.xlsx", "SVO", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E6").Value = "+" Then Call Perekhid("Perekhid.xlsx", "Perekhid", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E9").Value = "+" Then Call Luvers("Luvers.xlsx", "Luvers", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E7").Value = "+" Then Call Dovira("Dovira.xlsx", "Dovira_price.xlsx", "Dovira", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E13").Value = "+" Then Call RTM("RTM.xlsx", "RTM", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E18").Value = "+" Then Call Tristar("Tristar.xlsx", "Tristar_GRP.xlsx", "Tristar", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E14").Value = "+" Then Call Sean("Sean_city.xlsx", "Sean_board.xlsx", "Sean", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E11").Value = "+" Then Call Mallis("Mallis.xlsx", "Mallis_GRP.xlsx", "Mallis", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E10").Value = "+" Then Call Alhor("Alhor.xlsx", "Alhor", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E17").Value = "+" Then Call CityDnepr("CityDnepr.xlsx", "CityDnepr", "3.0�6.0", "1.2�1.8", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E19").Value = "+" Then Call Prospect("Prospect.xlsx", "Prospect_GRP.xlsx", "Prospect", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E12").Value = "+" Then Call Megapolis("Megapolis_Kh.xlsx", "Megapolis_Dp.xlsx", "Megapolis_GRP.xlsx", "Megapolis", "3�6", "1.2�1.8 2�3", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E15").Value = "+" Then Call Bomond("Bomond.xlsx", "Bomond_GRP.xlsx", "Bomond", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E20").Value = "+" Then Call ThreeSixDnepr("3x6Dnepr.xlsx", "3x6Dnepr", "Oblast", "Dnipro", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E21").Value = "+" Then Call NashaSprava("NashaSprava_board.xlsx", "NashaSprava", "�����", "NashaSprava_citylight.xlsx", "NashaSprava_scroll.xlsx", "NashaSprava_GRP.xlsx", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E22").Value = "+" Then Call MegapolisUA("Megapolis_UA.xlsx", "�������", "Megapolis", nameOfPathGeneralFile, nameOfGeneralFile)
    If ActiveWorkbook.Worksheets("������").Range("E23").Value = "+" Then Call T52("T52.xlsx", "����1", "T52", nameOfPathGeneralFile, nameOfGeneralFile)
    
    Call HomeAllSheets
    Windows(nameOfGeneralFile).Activate ' ��������� ����� ������ ������� ������
    Sheets("����� ����").Select
    Dim XCell As Object
    Dim XCol, XRow As Integer
    txtCol = "$$@@2"  ' ����� ��� �������
    Set XCell = ThisWorkbook.Sheets("����� ����").Cells.Find(txtCol)
    XCol = XCell.Column
    XRow = XCell.Row
'--------�����------------
    If ActiveWorkbook.Worksheets("������").Range("E3").Value = "+" Then
        Range("AL19:AM19").Select
        Selection.AutoFill Destination:=Range("AL19:" & "AM" & XRow - 2), Type:=xlFillDefault
        Range("AL20:" & "AM" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------��������------------
    If ActiveWorkbook.Worksheets("������").Range("E4").Value = "+" Then
        Range("AQ19:AR19").Select
        Selection.AutoFill Destination:=Range("AQ19:" & "AR" & XRow - 2), Type:=xlFillDefault
        Range("AQ20:" & "AR" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------���------------
    If ActiveWorkbook.Worksheets("������").Range("E13").Value = "+" Then
        Range("AV19:AW19").Select
        Selection.AutoFill Destination:=Range("AV19:" & "AW" & XRow - 2), Type:=xlFillDefault
        Range("AV20:" & "AW" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------�������------------
    If ActiveWorkbook.Worksheets("������").Range("E5").Value = "+" Then
        Range("BA19:BB19").Select
        Selection.AutoFill Destination:=Range("BA19:" & "BB" & XRow - 2), Type:=xlFillDefault
        Range("BA20:" & "BB" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------�������------------
    If ActiveWorkbook.Worksheets("������").Range("E6").Value = "+" Then
        Range("BF19:BG19").Select
        Selection.AutoFill Destination:=Range("BF19:" & "BG" & XRow - 2), Type:=xlFillDefault
        Range("BF20:" & "BG" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------������------------
    If ActiveWorkbook.Worksheets("������").Range("E7").Value = "+" Then
        Range("BK19:BL19").Select
        Selection.AutoFill Destination:=Range("BK19:" & "BL" & XRow - 2), Type:=xlFillDefault
        Range("BK20:" & "BL" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------�����------------
    If ActiveWorkbook.Worksheets("������").Range("E8").Value = "+" Then
        Range("BP19:BQ19").Select
        Selection.AutoFill Destination:=Range("BP19:" & "BQ" & XRow - 2), Type:=xlFillDefault
        Range("BP20:" & "BQ" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------������------------
    If ActiveWorkbook.Worksheets("������").Range("E9").Value = "+" Then
        Range("BV19:BW19").Select
        Selection.AutoFill Destination:=Range("BV19:" & "BW" & XRow - 2), Type:=xlFillDefault
        Range("BV20:" & "BW" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------������------------
    If ActiveWorkbook.Worksheets("������").Range("E10").Value = "+" Then
        Range("CA19:CB19").Select
        Selection.AutoFill Destination:=Range("CA19:" & "CB" & XRow - 2), Type:=xlFillDefault
        Range("CA20:" & "CB" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------������------------
    If ActiveWorkbook.Worksheets("������").Range("E11").Value = "+" Then
        Range("CF19:CG19").Select
        Selection.AutoFill Destination:=Range("CF19:" & "CG" & XRow - 2), Type:=xlFillDefault
        Range("CF20:" & "CG" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------���������------------
    If ActiveWorkbook.Worksheets("������").Range("E12").Value = "+" Or ActiveWorkbook.Worksheets("������").Range("E22").Value = "+" Then
        Range("CL19:CM19").Select
        Selection.AutoFill Destination:=Range("CL19:" & "CM" & XRow - 2), Type:=xlFillDefault
        Range("CL20:" & "CM" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------����------------
    If ActiveWorkbook.Worksheets("������").Range("E14").Value = "+" Then
        Range("CR19:CS19").Select
        Selection.AutoFill Destination:=Range("CR19:" & "CS" & XRow - 2), Type:=xlFillDefault
        Range("CR20:" & "CS" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------������------------
    If ActiveWorkbook.Worksheets("������").Range("E15").Value = "+" Then
        Range("CW19:CX19").Select
        Selection.AutoFill Destination:=Range("CW19:" & "CX" & XRow - 2), Type:=xlFillDefault
        Range("CW20:" & "CX" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------������ �������------------
    If ActiveWorkbook.Worksheets("������").Range("E16").Value = "+" Then
        Range("DB19:DC19").Select
        Selection.AutoFill Destination:=Range("DB19:" & "DC" & XRow - 2), Type:=xlFillDefault
        Range("DB20:" & "DC" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------���������------------
    If ActiveWorkbook.Worksheets("������").Range("E17").Value = "+" Then
        Range("DH19:DI19").Select
        Selection.AutoFill Destination:=Range("DH19:" & "DI" & XRow - 2), Type:=xlFillDefault
        Range("DH20:" & "DI" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------�������------------
    If ActiveWorkbook.Worksheets("������").Range("E18").Value = "+" Then
        Range("DM19:DN19").Select
        Selection.AutoFill Destination:=Range("DM19:" & "DN" & XRow - 2), Type:=xlFillDefault
        Range("DM20:" & "DN" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------��������------------
    If ActiveWorkbook.Worksheets("������").Range("E19").Value = "+" Then
        Range("DR19:DS19").Select
        Selection.AutoFill Destination:=Range("DR19:" & "DS" & XRow - 2), Type:=xlFillDefault
        Range("DR20:" & "DS" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------3��6 �����------------
    If ActiveWorkbook.Worksheets("������").Range("E20").Value = "+" Then
        Range("DW19:DX19").Select
        Selection.AutoFill Destination:=Range("DW19:" & "DX" & XRow - 2), Type:=xlFillDefault
        Range("DW20:" & "DX" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
'--------T52------------
    If ActiveWorkbook.Worksheets("������").Range("E23").Value = "+" Then
        Range("EC19:ED19").Select
        Selection.AutoFill Destination:=Range("EC19:" & "ED" & XRow - 2), Type:=xlFillDefault
        Range("EC20:" & "ED" & XRow - 2).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If

    
    
    Call HidSheets
    Sheets("����� ����").Activate
    Range("a1").Select

    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Time finish makros: " & Format((Timer - iTimer) / 86400, "Long Time"), vbExclamation, "" ' ������ ���-���-���
    
End Sub
Private Sub HidSheets()
Dim wsh As Worksheet, NoHid, i As Long, j As Long
NoHid = Array("����� ����", "������", "GRP", "����_����")    '������ ��� ����� ����� ���������
For Each wsh In ThisWorkbook.Worksheets
    j = 0
    For i = 0 To UBound(NoHid)
        If wsh.Name <> NoHid(i) Then j = j + 1
    Next i
    If j > UBound(NoHid) Then wsh.Visible = False
Next wsh
End Sub



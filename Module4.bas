Attribute VB_Name = "Module4"
Sub forClient()
Dim ws As Worksheet
ThisWorkbook.Save
For Each ws In ActiveWorkbook.Sheets
ws.Visible = True
Next
Sheets("����� ����").Activate
If Range("C8") = "POSTERSCOPE UKRAINE" Then Call Client
If Range("C8") = "Dentsu media" Then Call Dentsu
If Range("C8") = "Carat Ukraine" Then Call Carat
If Range("C8") = "Vizeum" Then Call Vizeum
If Range("C8") = "Isobar Ukraine" Then Call Isobar
End Sub
Sub Client()
Dim i As Integer
Dim Flag As Boolean
Dim lLastRow As Long
Dim lLastCol As Long
Dim nameOfGeneralFile As String
Dim nameOfPathGeneralFile As String
nameOfPathGeneralFile = ActiveWorkbook.Path
nameOfGeneralFile = ActiveWorkbook.Name
Dim ws As Worksheet

filepath = ThisWorkbook.Path
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

' ---------- ������ ����������� ----------
Cells.ClearOutline
' ---------- ������� ������ ----------
ActiveSheet.Buttons.Delete
    
    '-------- ������ ������� --------
    txtCol = "$$@@6"  ' ����� ��� �������
    Set totalCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If totalCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� "
    Else
    totalCol = totalCell.Column
    totalRow = totalCell.Row
    End If

    txtCol = "$$@@4"  ' ����� ��� �������
    Set XCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "���-�� ���������� "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    With ActiveSheet
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (totalRow - 1) _
               ).Copy
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (totalRow - 1) _
               ).PasteSpecial Paste:=xlPasteValues
    End With
    End If
    
    txtCol = "$$@@5"  ' ����� ��� �������
    Set XCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Reach 18+ per month (daily frequency 1+) "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    With ActiveSheet
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (totalRow - 1) _
               ).Copy
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (totalRow - 1) _
               ).PasteSpecial Paste:=xlPasteValues
    End With
    End If
    
        ' ---------- ������ own-����� ----------
    With ActiveSheet
        .Range( _
                ReturnName(totalCol) & ":" & ReturnName(totalCol + 1000) _
               ).Delete Shift:=xlToLeft
    End With

    Rows(XRow).Clear
    
    txtCol = "��  ���� ������� ���"  ' ����� ��� �������
    Set XCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "��  ���� ������� ��� "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    With ActiveSheet
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (XRow) _
               ).Copy
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (XRow) _
               ).PasteSpecial Paste:=xlPasteValues
    End With
    End If
      
    ' ---------- cancel CP mode ----------
    Application.CutCopyMode = False
    
    ' ---------- ������� �������� ������ ----------
    Cells.Validation.Delete
    
    ActiveWindow.FreezePanes = False
    Cells(1, 1).Select
    
For Each ws In ActiveWorkbook.Sheets
ws.Visible = True
Next
' ---------- ������ ����� ----------
Call deleteSheet
' ---------- ������ ������� ----------
Delete_Macroses
ThisWorkbook.Sheets("����� ����").Activate

' ---------- save new file ----------
Flag = Save_File_As("")
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Private Function Save_File_As(Doctype As String) As Boolean
Dim sfn As String
sfn = ActiveWorkbook.Name

Dim sSuff$: sSuff = Format(Now, "yyyymmdd")
Dim startsSuff$: startsSuff = ThisWorkbook.Sheets("����� ����").Range("D9").Value
Dim finishsSuff$:  If ThisWorkbook.Sheets("����� ����").Range("E9").Value = "" Then finishsSuff = "" Else finishsSuff = "-" & ThisWorkbook.Sheets("����� ����").Range("E9").Value

ChDir (ThisWorkbook.Path)
ActiveWorkbook.SaveAs _
Filename:="MP" & sSuff & "_" & ThisWorkbook.Sheets("����� ����").Range("C6").Value & "_OOH_" & startsSuff & finishsSuff & ".xlsx", _
FileFormat:=xlOpenXMLWorkbook, _
Password:="", _
WriteResPassword:="", _
CreateBackup:=False
Save_File_As = True
End Function

Private Sub Delete_Macroses()
    Dim oVBComponent As Object, lCountLines As Long
    'check if project is protected
    If ActiveWorkbook.VBProject.Protection = 1 Then
        MsgBox "VBProject is protected." & vbCrLf & _
             "     Components will not be deleted.", vbExclamation, "Execution canceled"
        Exit Sub
    End If
 
    For Each oVBComponent In ActiveWorkbook.VBProject.VBComponents
        On Error Resume Next
        With oVBComponent
            Select Case .Type
            Case 1    'Modules
                .Collection.Remove oVBComponent
            Case 2    'Class' modules
                .Collection.Remove oVBComponent
            Case 3    'Forms
                .Collection.Remove oVBComponent
            Case 100    'CurrentBook, Sheets
                    lCountLines = .CodeModule.CountOfLines
                    .CodeModule.DeleteLines 1, lCountLines
            End Select
        End With
    Next
    Set oVBComponent = Nothing
End Sub

Private Sub deleteSheet()
Dim wsh As Worksheet, NoHid, i As Long, j As Long
NoHid = Array("����� ����")    '������� ��� ����� ����� ���������
For Each wsh In ThisWorkbook.Worksheets
    j = 0
    For i = 0 To UBound(NoHid)
        If wsh.Name <> NoHid(i) Then j = j + 1
    Next i
    If j > UBound(NoHid) Then wsh.Delete
Next wsh
End Sub
Sub Carat()
Dim ws As Worksheet
Dim rCell As Range

'----------������ ������--------------
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = "Century Gothic"
         End With
    Next ws
'-----------������ ����� �������-------------
ThisWorkbook.Sheets("����� ����").Activate
    Application.FindFormat.Interior.Color = RGB(57, 129, 136)
    Application.ReplaceFormat.Interior.Color = RGB(0, 162, 215)
    Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True

'-----------������ ����� ������-------------
Dim lLastRow As Long
lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
For Each cell In ActiveSheet.Range(Cells(1, 2), Cells(lLastRow, 14)) 'ActiveCell.CurrentRegion
    If cell.Borders(xlEdgeTop).LineStyle <> xlNone Then cell.Borders(xlEdgeTop).Color = RGB(0, 162, 215)
    If cell.Borders(xlEdgeBottom).LineStyle <> xlNone Then cell.Borders(xlEdgeBottom).Color = RGB(0, 162, 215)
    If cell.Borders(xlEdgeLeft).LineStyle <> xlNone Then cell.Borders(xlEdgeLeft).Color = RGB(0, 162, 215)
    If cell.Borders(xlEdgeRight).LineStyle <> xlNone Then cell.Borders(xlEdgeRight).Color = RGB(0, 162, 215)
    If cell.Font.Color = RGB(57, 129, 136) Then cell.Font.Color = RGB(0, 162, 215)
  Next cell
'-----------������ ��������-------------
ActiveSheet.Shapes.Range(Array("Picture 8")).Select
Selection.Delete
Sheets("Logo").Select
ActiveSheet.Shapes.Range(Array("Picture 1")).Select
Selection.Copy
ThisWorkbook.Sheets("����� ����").Activate
ActiveSheet.Paste
Selection.ShapeRange.IncrementLeft 1060
Selection.ShapeRange.IncrementTop -42


Call Client
End Sub
Sub Vizeum()
Dim ws As Worksheet
Dim rCell As Range

'----------������ ������--------------
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = "Arial"
         End With
    Next ws
'-----------������ ����� �������-------------
ThisWorkbook.Sheets("����� ����").Activate
    Application.FindFormat.Interior.Color = RGB(57, 129, 136)
    Application.ReplaceFormat.Interior.Color = RGB(255, 192, 0)
    Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
'-----------������ ����� ������-------------
Dim lLastRow As Long
lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
For Each cell In ActiveSheet.Range(Cells(1, 2), Cells(lLastRow, 14)) 'ActiveCell.CurrentRegion
    If cell.Font.Color = RGB(57, 129, 136) Then cell.Font.Color = RGB(255, 192, 0)
  Next cell
'-----------������ ��������-------------
ActiveSheet.Shapes.Range(Array("Picture 8")).Select
Selection.Delete
Sheets("Logo").Select
ActiveSheet.Shapes.Range(Array("Picture 4")).Select
Selection.Copy
ThisWorkbook.Sheets("����� ����").Activate
ActiveSheet.Paste
Selection.ShapeRange.IncrementLeft 1010
Selection.ShapeRange.IncrementTop -42


Call Client
End Sub
Sub Isobar()
Dim ws As Worksheet
Dim rCell As Range

'----------������ ������--------------
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = "Arial"
         End With
    Next ws
'-----------������ ����� �������-------------
ThisWorkbook.Sheets("����� ����").Activate
    Application.FindFormat.Interior.Color = RGB(57, 129, 136)
    Application.ReplaceFormat.Interior.Color = RGB(249, 76, 7)
    Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True
'-----------������ ����� ������-------------
Dim lLastRow As Long
lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
For Each cell In ActiveSheet.Range(Cells(1, 2), Cells(lLastRow, 14)) 'ActiveCell.CurrentRegion
    If cell.Font.Color = RGB(57, 129, 136) Then cell.Font.Color = RGB(249, 76, 7)
  Next cell
'-----------������ ��������-------------
ActiveSheet.Shapes.Range(Array("Picture 8")).Select
Selection.Delete

Call Client
End Sub
Sub Dentsu()
Dim ws As Worksheet
Dim rCell As Range

'----------������ ������--------------
    For Each ws In Worksheets
         With ws
            .Cells.Font.Name = "Century Gothic"
         End With
    Next ws
'-----------������ ����� �������-------------
ThisWorkbook.Sheets("����� ����").Activate
    Application.FindFormat.Interior.Color = RGB(57, 129, 136)
    Application.ReplaceFormat.Interior.Color = RGB(89, 89, 89)
    Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=True, ReplaceFormat:=True

'-----------������ ����� ������-------------
Dim lLastRow As Long
lLastRow = Cells(Rows.Count, 2).End(xlUp).Row
For Each cell In ActiveSheet.Range(Cells(1, 2), Cells(lLastRow, 14)) 'ActiveCell.CurrentRegion
    If cell.Font.Color = RGB(57, 129, 136) Then cell.Font.Color = RGB(89, 89, 89)
  Next cell
'-----------������ ��������-------------
ActiveSheet.Shapes.Range(Array("Picture 8")).Select
Selection.Delete
Sheets("Logo").Select
ActiveSheet.Shapes.Range(Array("Picture 3")).Select
Selection.Copy
ThisWorkbook.Sheets("����� ����").Activate
ActiveSheet.Paste
Selection.ShapeRange.IncrementLeft 1000
Selection.ShapeRange.IncrementTop -42


Call Client
End Sub
Sub Buying()
Dim i As Integer
Dim Flag As Boolean
Dim lLastRow As Long
Dim lLastCol As Long
Dim nameOfGeneralFile As String
Dim nameOfPathGeneralFile As String
nameOfPathGeneralFile = ActiveWorkbook.Path
nameOfGeneralFile = ActiveWorkbook.Name
Dim ws As Worksheet

filepath = ThisWorkbook.Path
ThisWorkbook.Save
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

    
    '-------- ������ ������� --------
    txtCol = "$$@@6"  ' ����� ��� �������
    Set totalCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If totalCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "����� "
    Else
    totalCol = totalCell.Column
    totalRow = totalCell.Row
    End If
    
    txtCol = "$$@@5"  ' ����� ��� �������
    Set XCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    StrForMsgBox = StrForMsgBox + "Reach 18+ per month (daily frequency 1+) "
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    With ActiveSheet
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (totalRow - 1) _
               ).Copy
            .Range( _
                ReturnName(XCol) & (XRow) & ":" & ReturnName(XCol + 1) & (totalRow - 1) _
               ).PasteSpecial Paste:=xlPasteValues
    End With
    End If
    
    txtCol = "$$@@7"  ' ����� ��� �������
    Set XCell = ThisWorkbook.ActiveSheet.Cells.Find(txtCol)
    If XCell Is Nothing Then
    Else
    XCol = XCell.Column
    XRow = XCell.Row
    With ActiveSheet
            .Columns( _
                ReturnName(XCol) & ":" & ReturnName(XCol + 3) _
               ).Clear
            .Columns( _
                ReturnName(XCol) & ":" & ReturnName(XCol + 3) _
               ).Hidden = True
    End With
    End If

    ActiveSheet.Shapes.Range(Array("Button 80")).Select
    Selection.Delete
    ThisWorkbook.Sheets("����� ����").Cells(1, 1).Select
    Call HidSheets

' ---------- save new file ----------
Flag = Save_File_As_xlsm("")
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
Private Function Save_File_As_xlsm(Doctype As String) As Boolean
Dim sfn As String
sfn = ActiveWorkbook.Name

Dim sSuff$: sSuff = Format(Now, "hh-mm'ss''")
ChDir (ThisWorkbook.Path)
ActiveWorkbook.SaveAs _
Filename:="Posterscope_" & sfn & "_" & Doctype & "_Buying_" & sSuff & ".xlsm", _
FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
CreateBackup:=False
Save_File_As_xlsm = True
End Function
Private Sub HidSheets()
Dim wsh As Worksheet, NoHid, i As Long, j As Long
NoHid = Array("����� ����", "������", "Timeline")   '������ ��� ����� ����� ���������
For Each wsh In ThisWorkbook.Worksheets
    j = 0
    For i = 0 To UBound(NoHid)
        If wsh.Name <> NoHid(i) Then j = j + 1
    Next i
    If j > UBound(NoHid) Then wsh.Visible = False
Next wsh
End Sub

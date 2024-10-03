Attribute VB_Name = "Module2"
Function —÷≈ѕ»“№ƒ»јѕј«ќЌ(ƒиапазон—цепить As Range, ƒиапазон”словий As Range, ”словие As String, –азделитель As String)
    Dim i As Long
     
    'если диапазоны проверки и склеивани€ не равны друг другу - выходим с ошибкой
    If ƒиапазон”словий.Count <> ƒиапазон—цепить.Count Then
        —÷≈ѕ»“№ƒ»јѕј«ќЌ = CVErr(xlErrRef)
        Exit Function
    End If
     
    'проходим по все €чейкам, провер€ем условие и собираем текст в переменную OutText
    For i = 1 To ƒиапазон”словий.Cells.Count
        If ƒиапазон”словий.Cells(i) Like ”словие Then OutText = OutText & ƒиапазон—цепить.Cells(i) & –азделитель
    Next i
     
    'выводим результаты без последнего разделител€
    —÷≈ѕ»“№ƒ»јѕј«ќЌ = Left(OutText, Len(OutText) - Len(–азделитель))
End Function

Private Function Save_File_As(Doctype As String) As Boolean
Dim sfn As String ' им€ файла
Dim sSuff$: sSuff = Format(Now, "dd.mm.yyyy_hh-mm'ss''")
sfn = ActiveWorkbook.Name
ChDir (ThisWorkbook.Path & "\Vyborka")
ActiveWorkbook.SaveAs _
Filename:="Vyborka_" & sfn & "_" & Doctype & "_" & sSuff & ".xlsx", _
FileFormat:=xlOpenXMLWorkbook, _
Password:="", _
WriteResPassword:="", _
CreateBackup:=False
Save_File_As = True
End Function
Function ReturnName(ByVal num As Integer) As String
    ReturnName = Split(Cells(, num).Address, "$")(1)
End Function

Public Function NBUCURRENCY(currencyName As String, key As String, currencyDate As Date)
Dim sURI As String, oHttp As Object
    sURI = "https:" & Chr(47) & Chr(47) & "bank.gov.ua" & Chr(47) & "NBUStatService" & Chr(47) & "v1" & Chr(47) & "statdirectory" & Chr(47) & "exchange?valcode=" & currencyName & "&date=" & Format(currencyDate, "yyyymmdd") & "&json"
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    If oHttp Is Nothing Then Exit Function
    On Error GoTo ConnectionError
    oHttp.Open "GET", sURI, False
    On Error GoTo ConnectionError
    oHttp.Send
    NBUCURRENCY = jsonParse(oHttp.responseText, key)
ConnectionError:
    Set oHttp = Nothing
End Function
Private Function jsonParse(jsonStr As String, key As String)
    arr = Split(Replace(Replace(jsonStr, "[{", ""), "}]", ""), ",")
    For Each el In arr
        arr2 = Split(el, ":")
        arr2(0) = Replace(arr2(0), Chr(34), "")
        If arr2(0) = key Then
            If arr2(0) = "rate" Then jsonParse = CDbl(Replace(arr2(1), ".", ",")) Else: jsonParse = Replace(arr2(1), Chr(34), "")
            Exit For
        End If
    Next
End Function

Attribute VB_Name = "Module3"
Private Sub mReach()
    Application.ScreenUpdating = False 'отключение обновления экрана
    Workbooks.Application.DisplayAlerts = False ' отключение всплывающих окон

'On Error Resume Next

Excel.Sheets("mReach").Cells(4, 4) = ""
Excel.Sheets("mReach").Cells(5, 4) = ""

calc_type = Excel.Sheets("settings").Cells(1, 4)
dd = Excel.Sheets("mReach").Cells(5, 3)

If dd < 1 Or dd > 99 Then
    Excel.Sheets("mReach").Cells(5, 4) = "Количество дней от 1 до 99"
Else

If calc_type = 1 Then
    'GRP (расчет охватов)
    grp = Excel.Sheets("mReach").Cells(4, 3)
    dd = Excel.Sheets("mReach").Cells(5, 3)
    
    If grp < 10 Or grp > 120 Then
        Excel.Sheets("mReach").Cells(4, 4) = "GRP от 10 до 120"
    Else
    
    
    Excel.Sheets("mReach").Cells(9, 3) = "Охват @1+"
    Excel.Sheets("mReach").Cells(9, 4) = "Охват @3+"
    Excel.Sheets("mReach").Cells(9, 5) = "Охват @5+"
    
    grp1 = grp - grp Mod 5
    grp2 = grp1 + 5
    
    For i = 10 To 100
        city = Trim(Excel.Sheets("mReach").Cells(i, 2))
        If Len(city) = 0 Then Exit For
    
        Excel.Sheets("mReach").Cells(i, 3) = ""
        Excel.Sheets("mReach").Cells(i, 4) = ""
        Excel.Sheets("mReach").Cells(i, 5) = ""
    
    
    
        '@1+
        freq1 = Application.SumIfs(Range("Data!C11:C50104"), Range("Data!B11:B50104"), grp1, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
        freq2 = Application.SumIfs(Range("Data!C11:C50104"), Range("Data!B11:B50104"), grp2, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
    
        freq = freq1 + (freq2 - freq1) * (grp - grp1) / (grp2 - grp1)
        Excel.Sheets("mReach").Cells(i, 3) = Round(freq - 1.5, 0)
    
        '@3+
        freq1 = Application.SumIfs(Range("Data!D11:D50104"), Range("Data!B11:B50104"), grp1, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
        freq2 = Application.SumIfs(Range("Data!D11:D50104"), Range("Data!B11:B50104"), grp2, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
    
        freq = freq1 + (freq2 - freq1) * (grp - grp1) / (grp2 - grp1)
        Excel.Sheets("mReach").Cells(i, 4) = Round(freq - 1.5)
    
        '@5+
        freq1 = Application.SumIfs(Range("Data!E11:E50104"), Range("Data!B11:B50104"), grp1, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
        freq2 = Application.SumIfs(Range("Data!E11:E50104"), Range("Data!B11:B50104"), grp2, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
    
        freq = freq1 + (freq2 - freq1) * (grp - grp1) / (grp2 - grp1)
        Excel.Sheets("mReach").Cells(i, 5) = Round(freq - 1.5)
    
    
    Next i
    End If
End If



If calc_type = 2 Or calc_type = 3 Then
    '@1+ (расчет GRP)
    freq = Excel.Sheets("mReach").Cells(4, 3)
    dd = Excel.Sheets("mReach").Cells(5, 3)
    
    If freq < 0 Or freq > 91 Then
        Excel.Sheets("mReach").Cells(4, 4) = "Охват от 0 до 91"
    Else
    
    Excel.Sheets("mReach").Cells(9, 3) = "GRP"
    Excel.Sheets("mReach").Cells(9, 4) = ""
    Excel.Sheets("mReach").Cells(9, 5) = ""

    For i = 10 To 100
        city = Trim(Excel.Sheets("mReach").Cells(i, 2))
        If Len(city) = 0 Then Exit For
    
        Excel.Sheets("mReach").Cells(i, 3) = ""
        Excel.Sheets("mReach").Cells(i, 4) = ""
        Excel.Sheets("mReach").Cells(i, 5) = ""
        
        Excel.Sheets("mReach").Cells(i, 3).Select
        '=MATCH(FREQ;IF(Data!$F$11:$F$50140=DAYS;IF(Data!$A$11:$A$50140=CITY;Data!$C$11:$C$50140));1)
        If calc_type = 2 Then Selection.FormulaArray = "=MATCH(" & freq & ",IF(Data!R11C6:R50140C6=" & dd & ",IF(Data!R11C1:R50140C1=""" & city & """,Data!R11C3:R50140C3)),1)" 'FREQ1
        If calc_type = 3 Then Selection.FormulaArray = "=MATCH(" & freq & ",IF(Data!R11C6:R50140C6=" & dd & ",IF(Data!R11C1:R50140C1=""" & city & """,Data!R11C4:R50140C4)),1)" 'FREQ3
        
        idx = Excel.Sheets("mReach").Cells(i, 3)
        
        grp1 = Application.Index(Range("Data!B11:B50140"), idx)
        If IsError(grp1) = False Then
            grp2 = grp1 + 5
        
            If calc_type = 2 Then
                freq1 = Application.SumIfs(Range("Data!C11:C50104"), Range("Data!B11:B50104"), grp1, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
                freq2 = Application.SumIfs(Range("Data!C11:C50104"), Range("Data!B11:B50104"), grp2, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
            End If
            
            If calc_type = 3 Then
                freq1 = Application.SumIfs(Range("Data!D11:D50104"), Range("Data!B11:B50104"), grp1, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
                freq2 = Application.SumIfs(Range("Data!D11:D50104"), Range("Data!B11:B50104"), grp2, Range("Data!F11:F50104"), dd, Range("Data!A11:A50104"), city)
            End If
            
            grp = grp1 + (grp2 - grp1) * (freq - freq1) / (freq2 - freq1)
            Excel.Sheets("mReach").Cells(i, 3) = grp
        Else
                Excel.Sheets("mReach").Cells(i, 3) = "N/A"
        End If
        
        'Excel.Sheets("mReach").Cells(i, 4) = idx
        'Excel.Sheets("mReach").Cells(i, 5) = freq1
        'Excel.Sheets("mReach").Cells(i, 6) = freq2
        'Excel.Sheets("mReach").Cells(i, 7) = grp1
        'Excel.Sheets("mReach").Cells(i, 8) = grp2
        
    Next i
    End If
End If

End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub reach_total()
Dim i As Integer
ThisWorkbook.Sheets("mReach").Columns(9).ClearContents
    For i = 10 To 31 Step 1
        If ThisWorkbook.Sheets("mReach").Cells(i, 7).Value = "+" Then
            ThisWorkbook.Sheets("mReach").Cells(4, 3).Value = ThisWorkbook.Sheets("mReach").Cells(i, 8).Value
            Call mReach
            ThisWorkbook.Sheets("mReach").Cells(i, 9).Value = ThisWorkbook.Sheets("mReach").Cells(i, 3).Value / 100
        End If
    Next i
End Sub



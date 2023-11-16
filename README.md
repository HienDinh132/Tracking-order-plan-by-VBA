# Tracking-order-plan-by-VBA
Sub baocao()
    Dim banhang, mau, inbaocao, phu, final As Worksheet
    Dim dc_phu, i, dc_banhang, dc_banhang2, k, m, dc_banhang3, dc_banhang4, dc_inbaocao, dc_inbaocao1, j, dc_phu2, t, tong As Integer
    Dim twb As Workbook
    Dim path, fn As String
    
    
    Set twb = ThisWorkbook
    Set banhang = Sheets("Ban hang")
    Set mau = Sheets("mau")
    Set inbaocao = Sheets("in bao cao")
    Set phu = Sheets("Phu")
    Set final = Sheets("final")
    
    path = inbaocao.Range("H1")
    banhang.AutoFilterMode = False
    inbaocao.Rows("7:100000").Clear
    dc_banhang = banhang.Cells(Rows.Count, 1).End(xlUp).Row
    banhang.Range("A3:H" & dc_banhang).AutoFilter field:=6, _
            Criteria1:=">=" & inbaocao.Range("H2").Value, _
            Criteria2:="<=" & inbaocao.Range("H3").Value
    'banhang.Range("A3:N" & dc_banhang).SpecialCells(xlCellTypeVisible).Copy _
            phu.Range("A1")
    droplist
    
    dc_banhang2 = banhang.Cells(Rows.Count, 1).End(xlUp).Row
    phu.Cells.Clear
    banhang.Range("A3:H" & dc_banhang2).SpecialCells(xlCellTypeVisible).Copy _
            phu.Range("A1")
    dc_phu2 = phu.Cells(Rows.Count, 1).End(xlUp).Row
    
    tong = 0
    For t = 2 To dc_phu2
        If Len(phu.Range("A" & t).Value) > 0 Then
            tong = tong + 1
        End If
    Next t
    phu.Range("J" & 2).Value = tong
    j = dc_banhang2 + 1 - phu.Range("J" & 2).Value
    For i = j To dc_banhang2
        
        final.Cells.Clear
        inbaocao.Rows("7:100000").Clear
        
        inbaocao.Select
        inbaocao.Range("B5").NumberFormat = "mm-dd-yyyy"
        inbaocao.Range("B5").Value = banhang.Range("F" & i).Value
        For k = 1 To 5
            m = Application.WorksheetFunction.Match(inbaocao.Cells(6, k), _
                                                banhang.Range("A3:H3"), 0)
            inbaocao.Cells(7, k).Value = banhang.Cells(i, m).Value
        Next k
        
        dc_inbaocao = inbaocao.Cells(Rows.Count, 1).End(xlUp).Row
        mau.Range("A9:D13").Copy inbaocao.Range("A" & dc_inbaocao + 1)
        dc_inbaocao1 = inbaocao.Cells(Rows.Count, 1).End(xlUp).Row
        inbaocao.Range("A1:E" & dc_inbaocao1).SpecialCells(xlCellTypeVisible).Copy _
            final.Range("A1")
        final.Select
        final.Range("B5").ClearFormats
        final.Range("B5").NumberFormat = "mm-dd-yyyy"
        final.Range("D7").NumberFormat = "dd-mmm"
        'final.Range("B5").Value = Application.XLookup(final.Range("D7").Value, banhang.Range("G7:G10"), banhang.Range("F7:F10"))
        final.Copy
        ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path & banhang.Cells(i, 1) & ".pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
    Next i
    
End Sub

Sub droplist()
    Dim banhang, mau, inbaocao, phu As Worksheet
    Dim dc_phu, dc_banhang1, dc_phu2 As Integer
    
    Set banhang = Sheets("Ban hang")
    Set mau = Sheets("mau")
    Set inbaocao = Sheets("in bao cao")
    Set phu = Sheets("Phu")
    
    phu.Cells.Clear
    dc_banhang1 = banhang.Cells(Rows.Count, 1).End(xlUp).Row
    banhang.Range("F4:F" & dc_banhang1).Copy phu.Range("A1")
    dc_phu2 = phu.Cells(Rows.Count, 1).End(xlUp).Row
    phu.Range("B1").Value = dc_phu2
    phu.Columns(1).RemoveDuplicates Columns:=1, Header:=xlNo
    dc_phu = phu.Cells(Rows.Count, "A").End(xlUp).Row
    With inbaocao.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=Phu!A1:A" & dc_phu
    End With
End Sub

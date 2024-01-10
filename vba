Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

    Dim startDate As Date
    Dim Begin_Date As Date
    
    Dim diff As Long
    Dim minValue As Long
    Dim diff_display As Long
    
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim destSheet_copy As Worksheet
    
    
    Dim rng As Range
    Dim cell As Range
    
    Dim startRow As Long
    
    
    Dim sheet_3 As Worksheet
    Dim sheet_1 As Worksheet
    
    Dim endDate As Date
    
    
    Dim start_column_date As Integer
    Dim start_column_day As Integer
    
    Dim srcSheet_1 As Worksheet
    Dim srcSheet_3 As Worksheet
    Dim srcSheet_4 As Worksheet
    
    Dim rng_1 As Range
    Dim rng_3 As Range
    Dim rng_date As Range
    
    Dim rng_time As Range
    Dim targetCell As Range
    Dim shift As Integer
    
    
    Dim ws As Worksheet
    Dim chk As Shape
    Dim line1 As Shape
    Dim line2 As Shape
       
    Dim left As Double
    Dim top As Double
    
    Dim i As Integer
    Dim destRow As Integer
    Dim destColumn As Integer
    
    
    Dim j As Long
    
    Dim ostatok As Long
    Dim n As Long
    
    Dim min_minuts As Long
    'min_minuts = ThisWorkbook.Sheets(3).Cells(1, 2).Value
    'Debug.Print ("min_minuts : " & min_minuts)
   
    Dim Date_begin As Date
    Dim Search_Date As Date
    Dim Date_j As Date
    
Sub MainProcedure()

    ' Aucuaaai ia?ao? i?ioaao?o
    'Call CalculateTimeDifference

    ' Aucuaaai aoi?o? i?ioaao?o
    Call GroupRows
    Debug.Print ("groupRows")
    ' Aucuaaai o?aou? i?ioaao?o
    Call timeColumns

    ' Aucuaaai ?aoaa?oo? i?ioaao?o
    Call filling
    Debug.Print ("filling")

End Sub

Sub insert_value_cells_merge(i As Integer, destRow As Integer)
    With destSheet.Range("A" & destRow & ":AA" & destRow)
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = rng.Cells(i + 1, 2).Value
        .WrapText = True
        .Font.Size = 12
    '    .Name = "_" & rng.Cells(i + 1, 1).Value
    End With
    destSheet.Names.Add "_" & rng.Cells(i + 1, 1).Value, destSheet.Range("A" & destRow & ":AA" & destRow)
    'Debug.Print (" значение по имени ячейки : " & destSheet.Range("_" & rng.Cells(i + 1, 1).Value).Value)

End Sub

Sub insert_value_cells_merge_start_on(i As Integer, destRow As Integer, destColumn As Integer, destColumn_2 As Integer, sep As Integer)
    With destSheet.Range(destSheet.Cells(destRow, destColumn), destSheet.Cells(destRow, destColumn_2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 12
       ' Debug.Print (destRow)
       ' Debug.Print (destColumn)
       ' Debug.Print (sep)
        With destSheet.Range(destSheet.Cells(destRow, destColumn), destSheet.Cells(destRow, destColumn + sep))
        .Merge
        .Value = rng.Cells(i, 12).Value
        destSheet.Names.Add "_" & rng.Cells(i, 9).Value, destSheet.Range(destSheet.Cells(destRow, destColumn), destSheet.Cells(destRow, destColumn + sep))
        End With
        With destSheet.Range(destSheet.Cells(destRow, destColumn + sep + 1), destSheet.Cells(destRow, destColumn_2))
        .Merge
        destSheet.Names.Add "_2_" & rng.Cells(i, 9).Value, destSheet.Range(destSheet.Cells(destRow, destColumn + sep + 1), destSheet.Cells(destRow, destColumn_2))
        End With
    End With
    'destSheet.Names.Add "_" & rng.Cells(i, 9).Value, destSheet.Range(destSheet.Cells(destRow, destColumn), destSheet.Cells(destRow, destColumn_2))
End Sub

Sub insert_value_cells(i As Integer, destRow As Integer, destColumn As Integer)
    With destSheet.Cells(destRow, destColumn)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = rng.Cells(i, 12).Value
        .WrapText = True
        .Font.Size = 12
        '.Name = "_" & rng.Cells(i, 9).Value
    End With
    destSheet.Names.Add "_" & rng.Cells(i, 9).Value, destSheet.Cells(destRow, destColumn)
'    Debug.Print (" значение ячейки по имени : " & Range("_" & rng.Cells(i, 9).Value).Value)
End Sub

Sub insert_vved(destRow As Integer, vved As String, i As Integer)
    With destSheet.Range("A" & destRow & ":AA" & destRow)
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 10
        .Value = vved
        .Font.Bold = True
    End With
    'destSheet.Names.Add "_" & rng.Cells(i, 9).Value, destSheet.Range("A" & destRow & "AA" & destRow)
    'Debug.Print (" значение ячейки по имени группы : " & Range("_" & rng.Cells(i, 9).Value).Value)
End Sub

Function vvedenie(i As Integer) As String
    
    Dim rng_vved As Integer
    Dim vved As String
    rng_vved = rng.Cells(i, 26).Value
    
    Select Case rng_vved
        Case 80
            vved = "Гемокомпонент, внутривенно капельно"
        Case 81
            vved = "Внутривенно капельно"
        Case 82
            vved = "Внутривенно болюсно"
        Case 83
            vved = "Внутривенно микроструйно"
        Case 84
            vved = "Внутримышечно"
        Case 85
            vved = "Подкожно"
        Case 86
            vved = "Энтерально"
        Case 87
            vved = "Сублингвально"
        Case 88
            vved = "Суббуккально"
        Case 89
            vved = "Ректально"
        Case 91
            vved = "Эпидурально болюсно"
        Case 92
            vved = "Эпидурально микроструйно"
        Case 93
            vved = "Наружное применение"
        Case 94
            vved = "Ингаляционно"
        Case 95
            vved = "Назально"
        Case 96
            vved = "Глазные капли"
        Case 97
            vved = "Ушные капли"
        Case 98
            vved = "Ингаляционно длительно"
        Case 99
            vved = "Энтерально длительно"
        Case 111
            vved = "Энтерально через назогастральный зонд"
        Case 112
            vved = "Под коньюнктиву"
        Case 113
            vved = "Ретробульбарно"
        Case 114
            vved = "Интравитреально"
        Case 115
            vved = "Парабульбарно"
        Case 116
            vved = "Нижее веко"
        Case 237
            vved = "Через гастростому"
    End Select
    vvedenie = vved
End Function

Sub Naznachenie(i As Integer, destRow As Integer)
    'i_1 = i
    'destRow = destRow
    Dim vved As String
    
   ' Debug.Print (" i : " & i)
   ' Debug.Print (" destRow : " & destRow)
   ' Debug.Print (" Количество строк : " & rng.Rows.Count)
    For i = i To rng.Rows.Count
        vved = vvedenie(i)
      '  Debug.Print (" vved : " & vved)
      '  Debug.Print (" значение i: " & i)
        Call insert_vved(destRow, vved, i)
        destRow = destRow + 1
        While rng.Cells(i, 10).Value = rng.Cells(i + 1, 10).Value
            If IsEmpty(rng.Cells(i + 1, 1).Value) Then
                Exit For
            End If
    '        Debug.Print (" Значение ячейки 10 столбца : " & rng.Cells(i, 10).Value)
    '        Debug.Print (" Значение ячейки 10 столбца после " & rng.Cells(i + 1, 10).Value)
            Call insert_value_cells(i, destRow, 1)
            i = i + 1
            destRow = destRow + 1
        Wend
        Call insert_value_cells(i, destRow, 1)
        destRow = destRow + 1
    Next i
    
    
   ' Debug.Print (rng.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False))
   ' Debug.Print (rng.Cells(rng.Cells.Rows.Count, rng.Cells.Columns.Count).Address(RowAbsolute:=False, ColumnAbsolute:=False))
   ' Debug.Print (rng.Rows.Count + 10)
End Sub

Sub GroupRows()
    
    Set srcSheet = ThisWorkbook.Sheets(2) 'лист струтктуры
    Set destSheet = ThisWorkbook.Sheets(1) ' целевой лист
    Set destSheet_copy = ThisWorkbook.Sheets(4) 'сортировочный с листа 2
    
    srcSheet.UsedRange.Copy destSheet_copy.Cells(1, 1)
    Set rng = destSheet_copy.Range("A10:A" & srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row)
    
    With destSheet_copy.Sort
        .SortFields.Add key:=destSheet_copy.Range("I10"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange destSheet_copy.Range("A10:AY" & rng.Rows.Count + 10)
        .Header = xlNo
        .Apply
    End With
    
    Set rng = destSheet_copy.Range("A10:A" & srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row)

    startRow = 10
    destRow = 5
    
    
    For i = 1 To rng.Rows.Count
      '  Debug.Print (" destRow " & destRow)
      '  Debug.Print (" i " & i)
        'Назначения
        
        'События
        If rng.Cells(i, 2).Value = "События" Then
            'печать сноски событие
            Call insert_value_cells_merge(i, destRow)
            destRow = destRow + 1
            'вызов процедуры событий
            Call Events(i, destRow, "События")
        End If
        
        If rng.Cells(i, 2).Value = "Показатели" Then
            Call insert_value_cells_merge(i, destRow)
            destRow = destRow + 1
            'вызов процедуры событий
            Call Events(i, destRow, "Показатели")
        End If
        
        If rng.Cells(i + 1, 2).Value = "Манипуляции" Then
            Call insert_value_cells_merge(i, destRow)
            destRow = destRow + 1
            Call Events(i, destRow, "Манипуляции")
        End If
        
        If rng.Cells(i + 1, 2).Value = "Назначения" Then
            Call insert_value_cells_merge(i, destRow)
            destRow = destRow + 1
            Call Naznachenie(i, destRow)
        End If
        
        
        Call insert_value_cells_merge(i, destRow)
        destRow = destRow + 1
        While rng.Cells(i, 2).Value = rng.Cells(i + 1, 2).Value
     '       Debug.Print ("destRow : " & destRow)
     '       Debug.Print (" i " & i)
            Call insert_value_cells(i, destRow, 1)
            i = i + 1
            destRow = destRow + 1
        Wend
        Call insert_value_cells(i, destRow, 1)
        destRow = destRow + 1
    Next i
    'Debug.Print (" i : " & i)
    'Debug.Print (" destRow : " & destRow)
    'Debug.Print (" значение предпоследней ячейки : " & rng.Cells(i - 1, 2).Value)
        
    destRow = destRow + 1
   ' Debug.Print (" i : " & i)
   ' Debug.Print (" destRow : " & destRow)
   ' Debug.Print (" значение предпоследней ячейки " & rng.Cells(i - 1, 2).Value)
    destSheet.Cells(destRow + i, "A").Value = rng.Cells(i, 9).Value
    i = i + 1
    destRow = destRow + 1
    destSheet.Columns(2).AutoFit

End Sub

Sub Events(i As Integer, destRow As Integer, Value_string As String)
    
    Dim width_filling As Integer
    Dim pr_begin As Integer
    Dim res_cell As Variant
    Dim diff_res As Integer
    
    pr_begin = 0
    
    For i = i To rng.Rows.Count
        'Debug.Print (rng.Cells(i, 2).Value)
        If rng.Cells(i, 2).Value <> Value_string Then
            Exit Sub
        End If
        
        'Debug.Print ("значение 12: " & rng.Cells(i, 12).Value)
        'Debug.Print ("значение 23: " & rng.Cells(i, 23).Value)
        'Debug.Print (destRow)
        'Debug.Print ("width_filling : " & width_filling)
        If rng.Cells(i, 23).Value = 100 Then
       '     Debug.Print ("с новой строки : " & rng.Cells(i, 23).Value)
       '     Debug.Print ("значение : " & rng.Cells(i, 12).Value)
            Call insert_value_cells(i, destRow, 1)
            destRow = destRow + 1
            width_filling = 0
       '     Debug.Print ("destRow : " & destRow)
        ElseIf rng.Cells(i, 23).Value = 0 Then
       '     Debug.Print ("с новой строки : " & rng.Cells(i, 23).Value)
       '     Debug.Print ("значение : " & rng.Cells(i, 12).Value)
            Call insert_value_cells(i, destRow, 1)
            Call insert_value_cells_merge_start_on(i, destRow, 1, 27, 2)
            destRow = destRow + 1
            width_filling = 0
       '     Debug.Print ("destRow : " & destRow)
        Else
            If pr_begin = 0 Then
      '          Debug.Print (rng.Cells(i, 12).Value)
      '          Debug.Print (destRow)
                width_filling = rng.Cells(i, 23).Value
      '          Debug.Print ("width_filling " & width_filling)
                res_cell = return_cells_join_answer(i)
                
     '           Debug.Print ("Обратка 1 : " & res_cell(0))
     '           Debug.Print (" Обратка 2 : " & res_cell(1))
     '           Debug.Print (" destColumn : " & destColumn)
                Call insert_value_cells_merge_start_on(i, destRow, 1, 1 + res_cell(0), res_cell(0) - res_cell(1) - 1)
                pr_begin = 1
                destColumn = res_cell(0) + 1
            Else
                Do While width_filling < 99
                    width_filling = width_filling + rng.Cells(i, 23).Value
   '                 Debug.Print ("width : " & width_filling)
   '                 Debug.Print (rng.Cells(i, 12).Value)
                    res_cell = return_cells_join_answer(i)
   '                 Debug.Print (" Обратка 1 : " & res_cell(0))
   '                 Debug.Print (" Обратка 2 : " & res_cell(1))
   '                 Debug.Print (" destColumn : " & destColumn)
                    
    '                Debug.Print ("destRow : " & destRow)
                    
     '               Debug.Print (" destColumn : " & destColumn)
                    Call insert_value_cells_merge_start_on(i, destRow, destColumn + 1, destColumn + 1 + res_cell(0), res_cell(0) - res_cell(1) - 1)
                    i = i + 1
                    destColumn = destColumn + res_cell(0) + 1
                Loop
                destRow = destRow + 1
                i = i - 1
                pr_begin = 0
            End If
        End If
    Next i
End Sub

Sub timeColumns()
    Dim i As Integer
    Dim z As Integer
   
    Dim ws As Worksheet
    Set sheet_1 = ThisWorkbook.Sheets(1)
    i = 8
    While i <= 23
        For z = 2 To 17
            With sheet_1.Cells(4, z)
   '              Debug.Print (" z : " & z)
                 .NumberFormat = "General"
                 .HorizontalAlignment = xlCenter
                 .VerticalAlignment = xlCenter
                 .WrapText = True
                 .Value = "'" & Format(i, "00")
                 .Font.Size = 12
                 .Font.Color = RGB(128, 128, 128)
            End With
            i = i + 1
   '         Debug.Print (" i : " & i)
        Next z
    Wend
    i = 0
    While i <= 7
        For z = 18 To 25
            With sheet_1.Cells(4, z)
                 Debug.Print (" z : " & z)
                 .NumberFormat = "General"
                 .HorizontalAlignment = xlCenter
                 .VerticalAlignment = xlCenter
                 .WrapText = True
                 .Value = "'" & Format(i, "00")
                 .Font.Size = 12
                 .Font.Color = RGB(128, 128, 128)
            End With
            i = i + 1
   '         Debug.Print (" i : " & i)
        Next z
    Wend
    sheet_1.Cells(4, 26).Value = "C"
 End Sub
 
Function return_cells_join_answer(i As Integer) As Variant
    
    Dim cell_width As Double
    Dim cell_ost_width As Double
    
    cell_width = 24 / 100 * rng.Cells(i, 23).Value
    cell_ost_width = cell_width - Int(cell_width)
    
    'Debug.Print ("Остаток : " & cell_ost_width)
    
    If cell_ost_width >= 0.5 Then
        cell_width = WorksheetFunction.RoundUp(cell_width, 0)
    Else
        cell_width = WorksheetFunction.RoundDown(cell_width, 0)
    End If
    
    'Debug.Print ("результат округления : " & cell_width)
    'Debug.Print (" Результат функции : " & Int(cell_width / 2))
    return_cells_join_answer = Array(Int(cell_width), Int(cell_width / 2))
    
End Function
Function return_rgb_cell(hexcolor As String) As Variant
    
    Dim Red As String
    Dim Green As String
    Dim Blue As String
    
    Red = Mid((hexcolor), 1, 2)
    Green = Mid((hexcolor), 3, 2)
    Blue = Mid((hexcolor), 5, 2)
    
    'Debug.Print (Red)
    'Debug.Print (Green)
    'Debug.Print (Blue)
    
    return_rgb_cell = RGB(Application.WorksheetFunction.Hex2Dec(Red), Application.WorksheetFunction.Hex2Dec(Green), Application.WorksheetFunction.Hex2Dec(Blue))

End Function


Sub ClearSheet()
    Sheets("Лист1").Cells.Clear
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next shp

End Sub
Sub filling_naznachenie(valueToFind As String, rng_3_naznachenie As Range, rng_4_naznachenie As Range)
    
    Set destSheet = ThisWorkbook.Sheets(1)
    
    Dim minutes_nazn As Integer
    Dim diff_nazn As Integer
    Dim shp As Shape
    Dim rng_rec As Range
    Dim width_rec As Double
    
    Dim cell_3_row As Integer
    Dim cell_4_row As Integer
    Dim nm As Name
    Dim nm_dest As Variant
    Dim names_destSheet As Collection
    Set names_destSheet = New Collection
    
    For Each nm In destSheet.Names
        'Debug.Print (nm.Name)
        If nm.Name <> "Лист1!_" And Not nm.Name Like "Лист1!_2_*" And Not nm.Name Like "_2_*" And Not nm.Name Like "Лист1!_69*" Then
          '  Debug.Print (nm.Name)
            
            names_destSheet.Add nm.Name
        End If
        Next nm
    
    For Each nm_dest In names_destSheet
        'Debug.Print (nm_dest)
        For cell_4_row = 1 To rng_4_naznachenie.Rows.Count
            For cell_3_row = 1 To rng_3_naznachenie.Rows.Count
         '       Debug.Print ("3: " & rng_3_naznachenie.Cells(cell_3_row, 1).Value)
         '       Debug.Print ("nm.Name : " & nm_dest)
         '       Debug.Print (rng_3_naznachenie(cell_3_row, 1).Value)
         '       Debug.Print (rng_4_naznachenie(cell_4_row, 9).Value)
                If rng_3_naznachenie(cell_3_row, 1).Value = rng_4_naznachenie(cell_4_row, 9).Value And InStr(nm_dest, rng_4_naznachenie(cell_4_row, 9).Value) Then
                    width_rec = rng_3_naznachenie.Cells(cell_3_row, 8).Value / 60
                    minutes_nazn = rng_3_naznachenie.Cells(cell_3_row, 7).Value / 60
                    
         '           Debug.Print (" длина : " & width_rec)
         '           Debug.Print (" откуда : " & minutes_nazn)
         '           Debug.Print (rng_3_naznachenie.Cells(cell_3_row, 7).Value)
                    
                    Set rng_rec = Range(destSheet.Cells(destSheet.Range(nm_dest).Row, minutes_nazn + 2), destSheet.Cells(destSheet.Range(nm_dest).Row, minutes_nazn + 1 + width_rec))
                    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, rng_rec.left, rng_rec.top, rng_rec.Width, rng_rec.Height)
                    Debug.Print ("рисуем")
                    shp.TextFrame.Characters.Text = rng_3_naznachenie.Cells(cell_3_row, 5).Value & " " & rng_4_naznachenie.Cells(cell_4_row, 33).Value
                    shp.Fill.ForeColor.RGB = return_rgb_cell(rng_4_naznachenie.Cells(cell_4_row, 20).Value)
                   ' DoEvents
                   ' keycount_2 = keycount_2 + 1
                   ' keycount_2_2 = keycount_2
                   ' Debug.Print ("keycount_2 : " & keycount_2)
                End If
            Next cell_3_row
   '         keycount_1 = keycount_1 + 1
   '         keycount_1_1 = keycount_1
   '         Debug.Print ("конец внутреннего цикла, keycount_1_1 : " & keycount_1_1)
        Next cell_4_row
   Next nm_dest
Debug.Print ("here")
End Sub


Sub filling()
    
    Set srcSheet_3 = ThisWorkbook.Sheets(3)
    Set destSheet = ThisWorkbook.Sheets(1) ' целевой лист
    Set srcSheet_4 = ThisWorkbook.Sheets(4)
    
    Set rng_1 = destSheet.Range("A6:A" & destSheet.Cells(destSheet.Rows.Count, "A").End(xlUp).Row)
    
    Set rng_3 = srcSheet_3.Range("A5:A" & srcSheet_3.Cells(srcSheet_3.Rows.Count, "A").End(xlUp).Row)
    Set rng = srcSheet_4.Range("A10:A" & srcSheet_4.Cells(srcSheet_4.Rows.Count, "A").End(xlUp).Row)
    
    Dim cell As Range
    Dim cell_1 As Range
    Dim cell_2 As Range
    
    
    Dim foundCell As Range
    Dim valueToFind As String
    
    Dim nm As Name
    Dim minutes As Integer
    Dim diff As Integer
    Dim join_cell As Integer
    
    Dim pr_zap_2 As Integer
    Dim pr_zap_1 As Integer
    
    For Each nm In destSheet.Names
        'Debug.Print (" смотрим имя в книге: " & nm.Name)
        pr_zap_2 = 0
        For Each cell In rng_3
            valueToFind = cell.Value
            'Debug.Print ("Значение в листе 3(данные) : " & valueToFind)
            'Debug.Print ("Имя : " & nm.Name)
            If nm.Name = "Лист1!_" Or nm.Name = "_63333" Or nm.Name = "_63334" Or nm.Name = "_63335" Or nm.Name = "_63336" Or nm.Name = "_63337" Or nm.Name = "_63338" Then
                Exit For
            End If
            i = 1
            If InStr(nm.Name, valueToFind) > 0 Then
             '   Debug.Print ("значение которое найдено : " & valueToFind)
             '   Debug.Print ("имя ячейки : " & nm.Name)
                
                For i = i To rng.Rows.Count
             '       Debug.Print (" ячейка с 4 листа : " & rng.Cells(i, 9).Value)
             '       Debug.Print (" искомое значение  : " & valueToFind)
             '       Debug.Print (nm.Name)
                    If InStr(rng.Cells(i, 9).Value, valueToFind) > 0 Then
              '          Debug.Print (" значение разделенной строки  " & rng.Cells(i, 23).Value)
                        If rng.Cells(i, 23).Value = 100 Then
                            If rng_3.Cells(cell.Row - 4, 3).Value = "Назначения" Then
              '                  Debug.Print ("от листа 3 значение cell.Value, есть valueToFind : " & cell.Value)
              '                  Debug.Print (rng.Cells(rng.Rows.Count, 12).Value)
              '                  Debug.Print (" имя в листе имен : " & nm.Name)
              '                  Debug.Print (rng_3.Rows.Count)
              '                  Debug.Print (rng.Cells(rng.Rows.Count - 2, 9).Value)
               '                 Debug.Print (rng_3.Cells(rng_3.Rows.Count, 7))
               '                 Debug.Print (cell.Row)
               '                 Debug.Print (rng.Cells(i, 12).Value)
               '                 Debug.Print (rng.Cells(i, 9).Value)
               '                 Debug.Print (rng_3.Columns.Count)
               '                 Debug.Print (rng_3.Cells.Rows.Count)
               '                 Debug.Print ("от листа 4 значение ValuetoFind найденное : " & rng.Cells(i, 9).Value)
               '                 Debug.Print (" само значение ValueToFind : " & valueToFind)
               '
                                Dim rng_4_naznachenie As Range
                                Dim rng_3_naznachenie As Range
                                Set rng_4_naznachenie = Range(rng.Cells(i, 1), rng.Cells(rng.Rows.Count, 48))
                                
                                Set rng_3_naznachenie = Range(rng_3.Cells(cell.Row - 4, 1), rng_3.Cells(rng_3.Rows.Count, 18))
                '                Debug.Print (nm.Name)
                                'Dim Name_destSheet As String
                                'Name_destSheet = nm.Name
                                Call filling_naznachenie(valueToFind, rng_3_naznachenie, rng_4_naznachenie)
                                Exit Sub
                            End If
                            
                '            Debug.Print (valueToFind)
                '            Debug.Print (rng.Cells(i, 20).Value)
                '            Debug.Print (rng_3.Cells(cell.Row - 4, 8).Value)
                            
                            minutes = srcSheet_3.Cells(cell.Row, 7).Value
                '            Debug.Print (" minutes : " & minutes)
                            diff = Int(minutes / 60)
                '            Debug.Print ("diff : " & diff)
                            
                            If rng_3.Cells(cell.Row - 4, 8) = 60 Then
                                destSheet.Cells(destSheet.Range(nm.Name).Row, diff + 2).Value = srcSheet_3.Cells(cell.Row, 5).Value
                                destSheet.Cells(destSheet.Range(nm.Name).Row, diff + 2).Interior.Color = return_rgb_cell(rng.Cells(i, 20).Value)
                                
                            ElseIf rng_3.Cells(cell.Row - 4, 8) = 1 Then
                                destSheet.Cells(destSheet.Range(nm.Name).Row, diff + 2).Value = srcSheet_3.Cells(cell.Row, 5).Value
                                If rng.Cells(i, 12).Value = "Метки событий." Then
                                    destSheet.Cells(destSheet.Range(nm.Name).Row, diff + 2).Interior.Color = return_rgb_cell(rng.Cells(i, 20).Value)
                                End If
                            End If
                        
                        ElseIf rng.Cells(i, 23).Value = 0 Then
                                valueToFind = "_2_" & valueToFind
                                Debug.Print (" целая строка : " & rng.Cells(i, 23).Value)
                                If InStr(nm.Name, valueToFind) Then
                                    destSheet.Range(nm.Name).Value = srcSheet_3.Cells(cell.Row, 5).Value
                                    pr_zap_2 = 1
                                Exit For
                            End If
                        Else
                '            Debug.Print (" привет проценты : " & rng.Cells(i, 23).Value)
                '            Debug.Print (" привет проценты значение: " & rng.Cells(i, 23).Value)
                '            Debug.Print (" привет проценты адрес столбца: " & destSheet.Range(nm.Name).Columns.Address)
                '            Debug.Print (" привет проценты адрес строка : " & destSheet.Range(nm.Name).Rows.Address)
                '            Debug.Print (" имя в листе : " & nm.Name)
                '            Debug.Print (" привет проценты адрес строка : " & destSheet.Range(nm.Name).Rows.Address)
                            
                            valueToFind = "_2_" & valueToFind
                '            Debug.Print (valueToFind)
                            If InStr(nm.Name, valueToFind) Then
                                destSheet.Range(nm.Name).Value = srcSheet_3.Cells(cell.Row, 5).Value
                                pr_zap_2 = 1
                                Exit For
                            End If
                        End If
                    End If
                Next i
            End If
            If pr_zap_2 = 1 Then
                'Debug.Print ("Заполненность была, выходим")
                Exit For
            End If
        Next cell
    Next nm
End Sub













'подключение к dll
Declare PtrSafe Function GetBestStudent Lib "C:\Users\User\OneDrive\Desktop\Dll6\x64\Debug\Dll6.dll" _
    (ByRef marks As Double, _
     ByVal names_var As Variant, _
     ByVal rows As Long, _
     ByVal cols As Long, _
     ByRef the_best_cnt As Long _
    ) As Long

Const n_students As Long = 100
Const n_subjects As Long = 6

'основная процедура
Sub mainSub()
    SubjectsSub
    StudentFromAPI
    MarksSub
    ShowResults
End Sub


'процедура заполнения предметов
Sub SubjectsSub()
    Dim arr(1 To 6) As Variant
    arr(1) = "Математика"
    arr(2) = "Русский язык"
    arr(3) = "География"
    arr(4) = "Английский язык"
    arr(5) = "ОБЖ"
    arr(6) = "Физика"
    Dim j As Long
    For j = 1 To n_subjects
        Cells(1, j + 1).Value = arr(j)
    Next j
End Sub

'процедура заполнения учеников из файла на гитхабе какого-то человека, которого нашла в интернете
Function GetRussianNamesFromGitHub(cnt As Long) As Variant

    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url_names As String
    Dim url_surnames As String
    
    url_names = "https://raw.githubusercontent.com/Raven-SL/ru-pnames-list/refs/heads/master/lists/male_names_rus.txt"
    url_surnames = "https://raw.githubusercontent.com/Raven-SL/ru-pnames-list/refs/heads/master/lists/male_surnames_rus.txt"
    
    http.Open "GET", url_names, False: http.send
    If http.Status <> 200 Then Err.Raise 1001, , "Не удалось получить данные с names.txt"
    Dim names_text As String: names_text = http.responseText
    
    http.Open "GET", url_surnames, False: http.send
    If http.Status <> 200 Then Err.Raise 1002, , "Не удалось получить данные с surnames.txt"
    Dim surnames_text As String: surnames_text = http.responseText
    
    Dim name_arr As Variant: name_arr = Split(names_text, vbLf)
    Dim surname_arr As Variant: surname_arr = Split(surnames_text, vbLf)
    
    Dim result(): ReDim result(1 To cnt)
    Dim i As Long, ni As Long, si As Long
    Randomize
    For i = 1 To cnt
        ni = Int(Rnd * UBound(name_arr))
        si = Int(Rnd * UBound(surname_arr))
        result(i) = Trim(surname_arr(si)) & " " & Trim(name_arr(ni))
        
    Next i
    GetRussianNamesFromGitHub = result
End Function

'процедура для заполнения ячеек в таблице студентами, использующая функцию получения данных из гита
Sub StudentFromAPI()
    
    Dim list As Variant
    list = GetRussianNamesFromGitHub(n_students)
    
    Dim i As Long
    For i = 1 To n_students
        Cells(i + 1, 1).Value = list(i)
    Next i

End Sub
    

'процедура заполнения оценок с объектом Randomize(рнадомно выбираем оценки от 2 до 5)
Sub MarksSub()
    Dim r As Long, c As Long
    Randomize
    For r = 1 To n_students
        For c = 1 To n_subjects
        Cells(r + 1, c + 1).Value = Int(4 * Rnd + 2)
        Next c
    Next r
    
End Sub

'процедура для показания резульатов dll функции
Sub ShowResults()
    Dim src As Worksheet
    Set src = ThisWorkbook.Sheets("Лист1")
    
    Dim rng As Variant
    rng = src.Range("B2").Resize(n_students, n_subjects).Value
    
    Dim flat() As Double
    ReDim flat(0 To n_students * n_subjects - 1)
    
    Dim r As Long, c As Long, k As Long
    For r = 1 To n_students
        For c = 1 To n_subjects
            k = (r - 1) * n_subjects + (c - 1)
            flat(k) = CDbl(rng(r, c))
        Next c
    Next r

    Dim names_arr As Variant
    names_arr = Application.Transpose(src.Range("A2").Resize(n_students, 1).Value)
    
    Dim best_idx As Long, best_cnt As Long
    best_idx = GetBestStudent(flat(0), names_arr, n_students, n_subjects, best_cnt)
    
    Dim cnt_e As Long, cnt_g As Long, cnt_m As Long, cnt_b As Long
    Dim sum_all As Long, best_list As String
    Dim avg As Double, has_two As Boolean, scholarship As Long

    With src
        .Cells(1, 8).Value = "Средний"
        .Cells(1, 9).Value = "Стипендия"
        Const sb = 10000, sg = 7000, sm = 4000, sd = 0
        
        For r = 1 To n_students
            avg = 0: has_two = False
            For c = 1 To n_subjects
                If rng(r, c) = 2 Then has_two = True
                avg = avg + rng(r, c)
            Next c
            avg = avg / n_subjects

            If has_two Then
                scholarship = sd
                cnt_b = cnt_b + 1
                .Cells(r + 1, 8).Value = "Не аттестован"
             Else
                .Cells(r + 1, 8).Value = Round(avg, 2)
                Select Case True
                    Case avg >= 4.5
                        scholarship = 10000: cnt_e = cnt_e + 1
                    Case avg >= 4#
                        scholarship = 7000: cnt_g = cnt_g + 1
                    Case Else
                        stip = 4000: cnt_m = cnt_m + 1
                End Select
            End If
            .Cells(r + 1, 9).Value = scholarship
            sum_all = sum_all + scholarship
        Next r

        If best_idx = -1 Then
            .Cells(102, 1).Value = "Лучший: Нет отличников"
        Else
            .Cells(102, 1).Value = "Лучший: " & .Cells(bestIdx + 2, 1).Value
        End If
        .Cells(103, 1).Value = "Отличников: " & bestCnt
    End With

    best_list = ""
    If best_idx <> -1 Then
        Dim best_avg As Double
        best_avg = src.Cells(bestIdx + 2, 8).Value
        For r = 1 To n_students
            If src.Cells(r + 1, 8).Value = best_avg Then
                best_list = best_list & src.Cells(r + 1, 1).Value & ", "
            End If
        Next r
        If Len(best_list) > 2 Then best_list = Left(best_list, Len(best_list) - 2)
    Else
        best_list = "Нет отличников"
    End If

    Dim ws As Worksheet
    On Error Resume Next
      Set ws = ThisWorkbook.Sheets("Отчёт")
      If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Отчёт"
      Else
        ws.Cells.Clear
      End If
    On Error GoTo 0
    
    With ws
         .Range("A1").Value = "Сводный отчёт по студентам"
         .Range("A2").Value = "Лучшие студенты: " & best_list
         .Range("A3").Value = "Количество отличников: " & cnt_e
         .Range("A4").Value = "Количество хорошистов: " & cnt_g
         .Range("A5").Value = "Количество троечников: " & cnt_m
         .Range("A6").Value = "Количество двоечников: " & cnt_b
         .Range("A8").Value = "Общая сумма стипендий: " & summ_all & " руб"
         .Columns("A").AutoFit
    End With
    
    ImportReport "C:\1\TheBest.txt", "Отличники"
    ImportReport "C:\1\good.txt", "Хорошисты"
    ImportReport "C:\1\middle.txt", "Троечники"
    ImportReport "C:\1\TheWorst.txt", "Двоечники"
    
    ws.Activate
    MsgBox "Отчёты готовы: Сводка на листе отчёт, группы на отдельных вкладках.", vbInformation
    
End Sub

Sub ImportReport(filename As String, sheetName As String)
    Application.DisplayAlerts = False
    On Error Resume Next
      Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = sheetName

    With ws.QueryTables.Add(Connection:="TEXT;" & filename, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = True
        .TextFileDecimalSeparator = "."
        .TextFileThousandsSeparator = ","
        .Refresh BackgroundQuery:=False
    End With
    
    Dim last_row As Long, idx As Long, i As Long
    last_row = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow - 3
        If IsNumeric(ws.Cells(i, 1).Value) Then
            idx = CLng(ws.Cells(i, 1).Value)
            ws.Cells(i, 1).Value = src.Cells(idx + 2, 1).Value
        End If
    Next i
End Sub

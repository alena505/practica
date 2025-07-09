Option Explicit
'подключение к dll
Declare PtrSafe Function GetBestStudent Lib "C:\Users\User\OneDrive\Desktop\Dll6\x64\Debug\Dll6.dll" _
    (ByRef marks As Double, _
     ByVal names_var As Variant, _
     ByVal rows As Long, _
     ByVal cols As Long, _
     ByRef the_best_cnt As Long) As Long


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
    For j = 1 To 6
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
        ni = Int(Rnd * (UBound(name_arr) - 1) + 1)
        si = Int(Rnd * (UBound(surname_arr) - 1) + 1)
        result(i) = Trim(surname_arr(si)) & " " & Trim(name_arr(ni))
    Next i
    GetRussianNamesFromGitHub = result
End Function
'процедура для заполнения ячеек в таблице студентами, использующая функцию получения данных из гита
Sub StudentFromAPI()
    Dim list_names As Variant
    list_names = GetRussianNamesFromGitHub(100)
    Dim i As Long
    For i = 1 To 100
        Cells(i + 1, 1).Value = list_names(i)
    Next i
End Sub
'процедура заполнения оценок с объектом Randomize(рнадомно выбираем оценки от 2 до 5)
Sub MarksSub()
    Dim r As Long, c As Long
    Randomize
    For r = 2 To 101
        For c = 2 To 7
            Cells(r, c).Value = Int(4 * Rnd + 2)
        Next c
    Next r
End Sub
'процедура для показания резульатов dll функции
Sub ShowResults()
    Dim src As Worksheet: Set src = Worksheets("Лист1")
    src.Range("H2:I" & src.rows.Count).ClearContents

    Dim cnt_students As Long
    cnt_students = src.Cells(src.rows.Count, "A").End(xlUp).row - 1
    If cnt_students < 1 Then Exit Sub

    Dim rng As Variant
    rng = src.Range("B2").Resize(cnt_students, 6).Value

    Dim flat() As Double: ReDim flat(0 To cnt_students * 6 - 1)
    Dim r As Long, c As Long, k As Long
    For r = 1 To cnt_students
        For c = 1 To 6
            k = (r - 1) * 6 + (c - 1)
            flat(k) = CDbl(rng(r, c))
        Next c
    Next r

    Dim names_arr As Variant
    names_arr = Application.Transpose(src.Range("A2").Resize(cnt_students, 1).Value)

    Dim best_idx As Long, best_cnt As Long
    best_idx = GetBestStudent(flat(0), names_arr, cnt_students, 6, best_cnt)

    Dim cnt_e As Long, cnt_g As Long, cnt_m As Long, cnt_b As Long
    Dim sum_all As Long, avg As Double, has_two As Boolean, stp As Long
    Const SE As Long = 10000, SG As Long = 7000, SM As Long = 4000, SD As Long = 0

    src.Cells(1, 8).Value = "Средний балл"
    src.Cells(1, 9).Value = "Стипендия"

    For r = 1 To cnt_students
        avg = 0: has_two = False
        For c = 1 To 6
            If rng(r, c) = 2 Then has_two = True
            avg = avg + rng(r, c)
        Next c
        avg = avg / 6
        avg = Round(avg, 2)

        If has_two Then
            stp = SD: cnt_b = cnt_b + 1
            src.Cells(r + 1, 8).Value = avg
        Else
            src.Cells(r + 1, 8).Value = avg
            If avg >= 4.5 Then
                stp = SE: cnt_e = cnt_e + 1
            ElseIf avg >= 4 Then
                stp = SG: cnt_g = cnt_g + 1
            Else
                stp = SM: cnt_m = cnt_m + 1
            End If
        End If
        src.Cells(r + 1, 9).Value = stp
        sum_all = sum_all + stp
    Next r

    Dim ws As Worksheet
    On Error Resume Next: Set ws = Sheets("Отчёт"): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Отчёт"
    Else
        ws.Cells.Clear
    End If
    
    With ws
         .Range("A1").Value = "Сводный отчёт по студентам"
         
         If best_idx < 0 Then
            .Range("A2").Value = "Лучший студент: Нет данных"
         Else
            .Range("A2").Value = "Лучший студент: " & src.Cells(best_idx + 2, 1).Value
         End If
         
         .Range("A3").Value = "Количество отличников: " & cnt_e
         .Range("A4").Value = "Количество хорошистов: " & cnt_g
         .Range("A5").Value = "Количество троечников: " & cnt_m
         .Range("A6").Value = "Количество двоечников: " & cnt_b
         .Range("A8").Value = "Общая сумма стипендий: " & sum_all & " руб"
         .Columns("A").WrapText = True
         .rows.AutoFit
    End With
    
    CreateStudentSheet "Отличники", "TheBest.txt", src
    CreateStudentSheet "Хорошисты", "good.txt", src
    CreateStudentSheet "Троечники", "middle.txt", src
    CreateStudentSheet "Двоечники", "TheWorst.txt", src
    
    ws.Activate
    MsgBox "Отчёты готовы!", vbInformation
End Sub
'процедура Отчёта
Sub CreateStudentSheet(sheet_name As String, file_name As String, src As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(sheet_name).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = sheet_name
    
    ws.Cells(1, 1).Value = "№"
    ws.Cells(1, 2).Value = "ФИО"
    ws.Cells(1, 3).Value = "Средний балл"
    ws.Cells(1, 4).Value = "Стипендия"
    
    Dim fso As Object, ts As Object, line As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(file_name) Then Exit Sub
    Set ts = fso.OpenTextFile(file_name, 1, False)

    If Not ts.AtEndOfStream Then line = ts.ReadLine
    
    Dim row As Long: row = 2
    While Not ts.AtEndOfStream
        line = ts.ReadLine
        Dim parts As Variant
        parts = Split(line, vbTab)
        
        If UBound(parts) >= 3 Then
            Dim student_idx As Long
            student_idx = CLng(parts(0)) + 1
            
            ws.Cells(row, 1).Value = student_idx
            ws.Cells(row, 2).Value = src.Cells(student_idx, 1).Value
            ws.Cells(row, 3).Value = parts(2)
            ws.Cells(row, 4).Value = parts(3)
            row = row + 1
        End If
    Wend
    
    ts.Close
    
    With ws
        .Columns("A:D").AutoFit
        .Columns("B").WrapText = True
        .rows.AutoFit
    End With
End Sub










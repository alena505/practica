
Option Explicit

'подключение к dll
Declare PtrSafe Function GetBestStudent Lib "C:\Users\User\OneDrive\Desktop\Dll6\x64\Debug\Dll6.dll" _
    (ByRef marks As Double, ByVal rows As Long, ByVal cols As Long, ByRef TheBestCount As Long) As Long 'параметры, которые передаются через функцию на с++, возврат целочисленного значения(индекса лучшего ученика)

Const N_STUDENTS As Long = 100
Const N_SUBJECTS As Long = 6

'основная процедура
Sub mainSub()
    subjectsSub
    studentFromAPI
    marksSub
    ShowResults
End Sub


'процедура заполнения предметов
Sub subjectsSub()
    Dim arr(1 To 6) As Variant
    arr(1) = "Математика"
    arr(2) = "Русский язык"
    arr(3) = "География"
    arr(4) = "Английский язык"
    arr(5) = "ОБЖ"
    arr(6) = "Физика"
    Dim j As Long
    For j = 1 To N_SUBJECTS
        Cells(1, j + 1).Value = arr(j)
    Next j
End Sub

'процедура заполнения учеников из файла на гитхабе какого-то человека, которого нашла в интернете
Function GetRussianNamesFromGitHub(cnt As Long) As Variant

    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim urlNames As String
    Dim urlSurnames As String
    
    urlNames = "https://raw.githubusercontent.com/Raven-SL/ru-pnames-list/refs/heads/master/lists/male_names_rus.txt"
    urlSurnames = "https://raw.githubusercontent.com/Raven-SL/ru-pnames-list/refs/heads/master/lists/male_surnames_rus.txt"
    
    http.Open "GET", urlNames, False: http.send
    If http.Status <> 200 Then Err.Raise 1001, , "Не удалось получить данные с names.txt"
    Dim namesText As String: namesText = http.responseText
    
    http.Open "GET", urlSurnames, False: http.send
    If http.Status <> 200 Then Err.Raise 1002, , "Не удалось получить данные с surnames.txt"
    Dim surnamesText As String: surnamesText = http.responseText
    
    Dim nameArr As Variant: nameArr = Split(namesText, vbLf)
    Dim surnameArr As Variant: surnameArr = Split(surnamesText, vbLf)
    
    Dim result(): ReDim result(1 To cnt)
    Dim i As Long, ni As Long, si As Long
    Randomize
    For i = 1 To cnt
        ni = Int(Rnd * UBound(nameArr))
        si = Int(Rnd * UBound(surnameArr))
        result(i) = Trim(surnameArr(si)) & " " & Trim(nameArr(ni))
        
    Next i
    GetRussianNamesFromGitHub = result
End Function

'процедура для заполнения ячеек в таблице студентами, использующая функцию получения данных из гита
Sub studentFromAPI()
    
    Dim list As Variant
    list = GetRussianNamesFromGitHub(N_STUDENTS)
    
    Dim i As Long
    For i = 1 To N_STUDENTS
        Cells(i + 1, 1).Value = list(i)
    Next i

End Sub
    

'процедура заполнения оценок с объектом Randomize(рнадомно выбираем оценки от 2 до 5)
Sub marksSub()
    Dim r As Long, c As Long
    Randomize
    For r = 1 To N_STUDENTS
        For c = 1 To N_SUBJECTS
        Cells(r + 1, c + 1).Value = Int(4 * Rnd + 2)
        Next c
    Next r
    
End Sub

'процедура для показания резульатов dll функции
Sub ShowResults()
    Dim rng As Variant
    rng = Range("B2").Resize(N_STUDENTS, N_SUBJECTS).Value
    
    Dim flat() As Double
    ReDim flat(0 To N_STUDENTS * N_SUBJECTS - 1)
    Dim r As Long, c As Long, k As Long
    For r = 1 To N_STUDENTS
        For c = 1 To N_SUBJECTS
            k = (r - 1) * N_SUBJECTS + (c - 1)
            flat(k) = CDbl(rng(r, c))
        Next c
    Next r

    Dim bestIdx As Long, bestCnt As Long
    bestIdx = GetBestStudent(flat(0), N_STUDENTS, N_SUBJECTS, bestCnt)

    If bestIdx = -1 Then
        Cells(102, 1).Value = "Лучший: нет отличников"
    Else
        Cells(102, 1).Value = "Лучший: " & Cells(bestIdx + 2, 1).Value
    End If
    Cells(103, 1).Value = "Отличников: " & bestCnt

    Cells(1, 8).Value = "Средний балл"
    Cells(1, 9).Value = "Стипендия"
    
    Const STIP_BEST   As Long = 10000
    Const STIP_GOOD   As Long = 7000
    Const STIP_MIDDLE As Long = 4000
    Const STIP_BAD    As Long = 0
    
    Dim avg As Double, hasTwo As Boolean, stip As Long
    For r = 1 To N_STUDENTS
        avg = 0: hasTwo = False
        For c = 1 To N_SUBJECTS
            If rng(r, c) = 2 Then hasTwo = True
            avg = avg + rng(r, c)
        Next c
        avg = avg / N_SUBJECTS
        
        If hasTwo Then
            stip = STIP_BAD
            Cells(r + 1, 8).Value = "Не аттестованы"
        Else
            Cells(r + 1, 8).Value = Round(avg, 2)
            Select Case True
                Case avg >= 4.5: stip = STIP_BEST
                
                Case avg >= 4#:  stip = STIP_GOOD
                Case Else: stip = STIP_MIDDLE
            End Select
        End If
        Cells(r + 1, 9).Value = stip
    Next r
    
    If bestIdx <> -1 Then
        Dim bestAvg As Double
        bestAvg = Cells(bestIdx + 2, 8).Value
        
        Dim tiedNames As String: tiedNames = ""
        For r = 1 To N_STUDENTS
            If Cells(r + 1, 8).Value = bestAvg Then
                tiedNames = tiedNames & Cells(r + 1, 1).Value & ", "
            End If
        Next r
        
        If Len(tiedNames) >= 2 Then
            tiedNames = Left(tiedNames, Len(tiedNames) - 2)
        End If
    
        Cells(102, 1).Value = "Лучшие: " & tiedNames
        Else
            
    
    End If
            
End Sub
















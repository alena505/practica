Attribute VB_Name = "Module1"
Option Explicit

'подключение к dll
Declare PtrSafe Function GetBestStudent Lib "C:\Users\User\OneDrive\Desktop\Dll6\x64\Debug\Dll6.dll" _
    (ByRef marks As Double, ByVal rows As Long, ByVal cols As Long, ByRef TheBestCount As Long) As Long  'параметры, которые передаются через функцию на с++, возврат целочисленного значения(индекса лучшего ученика)

'основная процедура
Sub mainSub()
    subjectsSub
    studentsSub
    marksSub
    ShowResults
End Sub

'процедура заполнения предметов
Sub subjectsSub()
    Dim arr(1 To 6) As String
    arr(1) = "Математика"
    arr(2) = "Русский язык"
    arr(3) = "География"
    arr(4) = "Английский язык"
    arr(5) = "ОБЖ"
    arr(6) = "Физика"
    Dim i As Integer
    For i = 1 To 6
        Cells(1, i + 1).Value = arr(i)
    Next i
End Sub

'процедура заполнения учеников
Sub studentsSub()
    Dim arr2(1 To 20) As String
    arr2(1) = "Авдеева А.С."
    arr2(2) = "Семёнов А.М."
    arr2(3) = "Глазков П.В."
    arr2(4) = "Мухаяров В.А."
    arr2(5) = "Иванова Г.С."
    arr2(6) = "Иванов М.Ю."
    arr2(7) = "Сидоров К.В."
    arr2(8) = "Козлова С.Ф."
    arr2(9) = "Кирикоич П.Д."
    arr2(10) = "Дудник А.К."
    arr2(11) = "Мартынюк П.А."
    arr2(12) = "Леоньтев С.Б."
    arr2(13) = "Авдеев А.С."
    arr2(14) = "Кириллова К.О."
    arr2(15) = "Бондаренко А.В."
    arr2(16) = "Кушнаренко Е.В."
    arr2(17) = "Берёза Е.О."
    arr2(18) = "Гуреньтева С.Ю."
    arr2(19) = "Рудников К.В."
    arr2(20) = "Гаврилова М.Е."
    Dim i As Integer
    Randomize
    For i = 1 To 100
        Cells(i + 1, 1).Value = arr2(Int((20 * Rnd) + 1))
    Next i
End Sub

'процедура заполнения оценок с объектом Randomize(рнадомно выбираем оценки от 2 до 5)
Sub marksSub()
    Dim i As Integer
    Dim j As Integer
    Randomize
    For i = 1 To 100
        For j = 1 To 6
        Cells(i + 1, j + 1).Value = Int(4 * Rnd + 2)
        Next j
    Next i
    
End Sub

'процедура для показания резульатов dll функции
Sub ShowResults()
    Dim namesArr As Variant, tmp As Variant
    tmp = Range("B2:G101").Value
    namesArr = Application.Transpose(Range("A2:A101").Value)
    
    Dim rows As Long: rows = UBound(tmp, 1)
    Dim cols As Long: cols = UBound(tmp, 2)
    Dim flat() As Double
    ReDim flat(0 To rows * cols - 1)
    
    Dim r As Long, c As Long
    For r = 1 To rows
        For c = 1 To cols
            flat((r - 1) * cols + (c - 1)) = CDbl(tmp(r, c))
        Next c
    Next r

    
    Dim bestStudentIdx As Long
    Dim bestStudentsCount As Long
    bestStudentIdx = GetBestStudent(flat(0), rows, cols, bestStudentsCount)
    
    Cells(103, 1).Value = "Количество отличников: " & bestStudentsCount
    
    If bestStudentIdx = -1 Then
        Cells(102, 1).Value = "Лучший: никто (все «не аттестованы»)"
    Else
        Cells(102, 1).Value = "Лучший: " & namesArr(bestStudentIdx + 1)
    End If
    
    Cells(1, 8).Value = "Средний балл"
    Dim avg As Double, hasTwo As Boolean
    For r = 1 To rows
        avg = 0: hasTwo = False
        For c = 1 To cols
            If flat((r - 1) * cols + (c - 1)) = 2 Then hasTwo = True
            avg = avg + flat((r - 1) * cols + (c - 1))
        Next c
        If hasTwo Then
            Cells(r + 1, 8).Value = "не аттестован"
        Else
            Cells(r + 1, 8).Value = Round(avg / cols, 2)
        End If
    Next r
End Sub


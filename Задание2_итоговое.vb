Attribute VB_Name = "Module1"
Option Explicit

'подключение к dll
Declare PtrSafe Function GetBestStudent Lib "C:\Users\User\OneDrive\Desktop\Dll6\x64\Debug\Dll6.dll" _
    (ByRef marks As Double, ByVal rows As Long, ByVal cols As Long) As Long 'параметры, которые передаются через функцию на с++, возврат целочисленного значения(индекса лучшего ученика)

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
    Dim arr2(1 To 10) As String
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
    Dim i As Integer
    For i = 1 To 10
        Cells(i + 1, 1).Value = arr2(i)
    Next i
End Sub

'процедура заполнения оценок с объектом Randomize(рнадомно выбираем оценки от 2 до 5)
Sub marksSub()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 10
        For j = 1 To 6
        Randomize
        Cells(i + 1, j + 1).Value = Int(4 * Rnd + 2)
        Next j
    Next i
    
End Sub

'Процедура анализа результатов и вывода лучшего ученика
Sub ShowResults()
    Dim namesArr As Variant, tmp As Variant
    tmp = Range("B2:G11").Value
    namesArr = Application.Transpose(Range("A2:A11").Value)
    
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
    
    Dim bestIdx As Long
    bestIdx = GetBestStudent(flat(0), rows, cols)
    
    If bestIdx = -1 Then
        Cells(12, 1).Value = "Лучший: никто (все «не аттестованы»)"
    Else
        Cells(12, 1).Value = "Лучший: " & namesArr(bestIdx + 1)
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


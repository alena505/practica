Option Explicit

Sub mainSub()
    Call subjectsSub
    Call studentsSub
    Call marksSub
    
End Sub

Sub subjectsSub()
    Dim arr(1 To 6) As String
    arr(1) = "Математика"
    arr(2) = " Русский язык"
    arr(3) = "География"
    arr(4) = "Английский язык"
    arr(5) = "ОБЖ"
    arr(6) = "Физика"
    Dim i As Integer
    For i = 1 To 6
        Cells(1, i + 1).Value = arr(i)
    Next i
End Sub

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

Sub marksSub()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To 10
        For j = 1 To 6
        Randomize
        Cells(i + 1, j + 1).Value = Int((5 - 2 + 1) * Rnd + 2)
        Next j
    Next i
    
End Sub


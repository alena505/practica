Attribute VB_Name = "Module1"
Option Explicit

Declare PtrSafe Function coolSub Lib "C:\1\Dll1.dll" _
     (ByRef marks As Double, _
     ByRef names As LongPtr, _
     ByVal rows As Long, _
     ByVal cols As Long) As String
     
     
Sub mainSub()
    Call subjectsSub
    Call studentsSub
    Call marksSub
    Call cool2Sub
    
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



Sub cool2Sub()
    Dim namesRange As Range
    Dim marksRange As Range
    Dim namesArray() As String
    Dim marksArray() As Double
    Dim bestStudent As String
    Dim i As Long, j As Long
    
    
    Set namesRange = Range("A2:A11")
    Set marksRange = Range("B2:G11")
    
    
    ReDim namesArray(1 To namesRange.rows.Count)
    ReDim marksArray(1 To marksRange.rows.Count, 1 To marksRange.Columns.Count)
    
    For i = 1 To namesRange.rows.Count
        namesArray(i) = namesRange.Cells(i, 1).Value
    Next i
    
    For i = 1 To marksRange.rows.Count
        For j = 1 To marksRange.Columns.Count
            marksArray(i, j) = marksRange.Cells(i, j).Value
        Next j
    Next i
    
    bestStudent = coolSub(marksArray(1, 1), VarPtr(namesArray(1)), UBound(namesArray), 6)
    Range("A12").Value = "Лучший:" & bestStudent


End Sub




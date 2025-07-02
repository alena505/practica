VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub mainSub()
    Call subjectsSub
    Call studentsSub
    Call marksSub
    Call coolSub
    
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

Sub coolSub()
    Dim sum As Integer
    Dim hastwo As Boolean
    Dim count As Integer
    Dim sred As Double
    Dim bestsred As Double
    bestsred = 0
    Dim bestStudent As Integer
    bestStudent = 0
    Dim i As Integer
    Dim j As Integer
        
    Cells(1, 8).Value = "Сред.Балл"
    For i = 1 To 10
        sum = 0
        hastwo = False
        
        For j = 1 To 6
            If Cells(i + 1, j + 1).Value = 2 Then
                hastwo = True
                Cells(i + 1, j + 1).Interior.Color = RGB(255, 0, 0)
                Cells(i + 1, 8).Value = "Не аттестован"
                Cells(i + 1, 1).Interior.Color = RGB(255, 0, 0)
                
            Else
                sum = sum + Cells(i + 1, j + 1).Value
            End If
        Next j
        
    
        If Not hastwo Then
            sred = sum / 6
            Cells(i + 1, j + 1).Value = sred
            If sred > bestsred Then
                bestsred = sred
                bestStudent = i
            End If
        End If
    Next i
    
    
    If bestStudent > 0 Then
        Cells(15, 1).Value = "Лучший:"
        Cells(15, 2).Value = Cells(bestStudent + 1, 1).Value
        Cells(15, 3).Value = bestsred
        Cells(bestStudent + 1, 10).Value = "Лучший"
        Cells(bestStudent + 1, 10).Interior.Color = RGB(0, 255, 0)
        Cells(bestStudent + 1, 1).Interior.Color = RGB(0, 255, 0)
    Else
        Cells(15, 1).Value = "Нет подходящих студентов"
        Cells(15, 2).Value = ""
    End If
End Sub


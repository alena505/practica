Attribute VB_Name = "Module1"
Option Explicit

Public Type Batch
total As Long
brak As Long
crit As Long
name As LongPtr
End Type

'Подключение к dll
Declare PtrSafe Function AnalyzeProduction Lib "C:\Users\User\OneDrive\Desktop\Dll7\x64\Debug\Dll7.dll" _
    (ByVal parts As LongPtr, ByVal total As Long) As Long
    
Public Const fullN As Long = 100
Public testCounter As Long



'Основная функция
Sub main()
    Call AddString
    Call DetailsFromAPI
    Call MastersFromApi
    Call AddData
    Call Analyze
End Sub

'Процедура добавление наименования полей(строк)
Sub AddString()
    Dim w As Worksheet: Set w = ThisWorkbook.Worksheets("Лист1")
    w.Range("A1:E1").Value = Array("Наименование детали", "Количество произведённых", "Количество бракованных", "Категория детали", "Исполнитель(мастер)")
    w.Columns("A:E").AutoFit
End Sub

'Подключение(парсинг) к гиту(текстовому файлу), чтобы получить наименования деталей
Function GetDetailsFromGitHub(cnt As Long) As Variant

    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    
    
    url = "https://raw.githubusercontent.com/alena505/-/refs/heads/main/computer.txt?token=GHSAT0AAAAAADG37E4HJNQHAYKLXN4EKZZU2DRQJRA"
    
    
    http.Open "GET", url, False: http.send
    If http.Status <> 200 Then Err.Raise 1001, , "Не удалось получить данные с details.txt"
    Dim details_text As String: details_text = http.responseText
    
    
    Dim name_arr As Variant: name_arr = Split(details_text, vbLf)

    
    Dim result(): ReDim result(1 To cnt)
    Dim i As Long, si As Long
    Randomize
    For i = 1 To cnt
        si = Int(Rnd * UBound(name_arr))
        result(i) = Trim(name_arr(si))
        
    Next i
    GetDetailsFromGitHub = result
End Function
'Заполнение наименования деталек из гита
Sub DetailsFromAPI()
    
    Dim list As Variant
    Dim a As Worksheet: Set a = ThisWorkbook.Worksheets("Лист1")
    list = GetDetailsFromGitHub(100)
    
    a.Range("A2:A" & a.Rows.Count).ClearContents
    
    Dim i As Long
    For i = 1 To 100
        a.Cells(i + 1, 1).Value = list(i)
    Next i
    
    a.Columns("A").AutoFit

End Sub
'Подключение(парсинг) к гиту(текстовому файлу), чтобы получить имена мастеров
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

'Заполнение мастеров из гита
Sub MastersFromApi()
    Dim list_names As Variant
    list_names = GetRussianNamesFromGitHub(100)
    Dim i As Long
    For i = 1 To 100
        Cells(i + 1, 5).Value = list_names(i)
    Next i
End Sub
'Заполнение полей(строк) информации о деталях
Sub AddData()
    Randomize
    Dim i As Integer
    Dim total, brak As Integer
    For i = 1 To 100
        total = Int((20 * Rnd) + 1)
        brak = Int(total * Rnd)
        Cells(i + 1, 2).Value = total
        Cells(i + 1, 3).Value = brak
        Cells(i + 1, 4).Value = Int(2 * Rnd + 1)
    Next i
End Sub

'Процедура заполнения данных из dll
Sub Analyze()
    Dim list() As Long
    Dim i As Long
    Dim ta As Long
    Dim tb As Long
    Dim list_brak() As Long
    Dim b() As Byte
    Dim names() As LongPtr
    Dim s As String
    Dim j As Long
    Dim file_data As Integer
    Dim txt_line As String
    Dim count_bad As Long
    
    Dim all_data() As Batch
    
    
    
    Dim k As Long
    Dim summ As Long
   
    
   ' ReDim list_brak(1 To fullN) As Long
    'ReDim list_type(1 To fullN) As Long
    
    ReDim all_data(0 To fullN) As Batch
    
    
    
    'ReDim list(1 To fullN) As Long
    
    
    summ = 0
    For i = 1 To fullN
        summ = Len(Cells(i + 1, 5).Value) + summ + 1
    Next i
    
    
    
    Dim all_bytes() As Byte
    
    ReDim all_bytes(summ)
    
    j = 1
    For i = 1 To fullN
        'list(i) = CLng(Cells(i + 1, 2).Value)
        'list_brak(i) = CLng(Cells(i + 1, 3).Value)
        'list_type(i) = CLng(Cells(i + 1, 4).Value)
        
        all_data(i).total = CLng(Cells(i + 1, 2).Value)
        all_data(i).brak = CLng(Cells(i + 1, 3).Value)
        all_data(i).crit = CLng(Cells(i + 1, 4).Value)
        
        b = StrConv(Cells(i + 1, 5).Value, vbFromUnicode)
        
        
        
        For k = LBound(b) To UBound(b)
            all_bytes(j) = b(k)
            j = j + 1
        Next k
        
        all_bytes(j) = 0
        
        j = j + 1
        
        
        all_data(i).name = VarPtr(all_bytes(j - (UBound(b) - LBound(b) + 2)))
        'Cells(i + 1, 6).Value = CStr(Chr(b(0)))
    Next i
    
    
    
    
    
    count_bad = AnalyzeProduction(VarPtr(all_data(0)), fullN)
    Cells(103, 1).Value = "Число нарушителей: " & CStr(count_bad)
    
    
    testCounter = testCounter + 1
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:= _
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.name = "Мастера, сделавшие брак(Over)" & CStr(testCounter)
    
    file_data = FreeFile()
    Open "OverDefected.txt" For Input As #file_data
    i = 1
    While Not EOF(file_data)
        Line Input #file_data, txt_line
        ws.Cells(i, 1).Value = txt_line
        
        i = i + 1
    Wend
    
    If count_bad > 20 Then
        MsgBox ("Количество нарушителей слишком велико!")
    End If
    
    
    
    Set ws = ThisWorkbook.Sheets.Add(After:= _
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.name = "Анализ(Summary)" & CStr(testCounter)
    
    
    file_data = FreeFile()
    Open "ProductionSummary.txt" For Input As #file_data
    i = 1
    While Not EOF(file_data)
        Line Input #file_data, txt_line
        ws.Cells(i, 2).Value = txt_line
        
        i = i + 1
    Wend
    
    ws.Cells(1, 1).Value = "Полное количество деталей"
    ws.Cells(2, 1).Value = "Количество бракованных деталей"
    ws.Cells(3, 1).Value = "Процент брака"
    
    
    
    
End Sub








































































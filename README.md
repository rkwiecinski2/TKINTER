Option Explicit

Dim QuestionsDict As Object
Dim CurrentID As String
Dim OutputText As String

Private Sub UserForm_Initialize()
    Set QuestionsDict = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim qID As String, qText As String, qType As String
    Dim opt1 As String, opt2 As String, nextIfD As String, nextIfE As String
    Dim comment As String
    Dim q As Object
    
    Set ws = ThisWorkbook.Sheets("Questions") ' Upewnij się, że Twoja tabela jest na arkuszu o tej nazwie
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Wczytaj wszystkie pytania do słownika
    For i = 2 To lastRow
        qID = Trim(ws.Cells(i, 1).Value)
        If qID <> "" Then
            Set q = CreateObject("Scripting.Dictionary")
            q.Add "Question", ws.Cells(i, 2).Value
            q.Add "Type", ws.Cells(i, 3).Value
            q.Add "Option1", ws.Cells(i, 4).Value
            q.Add "Option2", ws.Cells(i, 5).Value
            q.Add "NextIfD", ws.Cells(i, 6).Value
            q.Add "NextIfE", ws.Cells(i, 7).Value
            q.Add "Comment", ws.Cells(i, 9).Value
            QuestionsDict.Add qID, q
        End If
    Next i

    CurrentID = "ID1"
    OutputText = ""
    ShowQuestion CurrentID
End Sub

Private Sub ShowQuestion(qID As String)
    Dim q As Object
    Set q = QuestionsDict(qID)
    
    lblQuestion.Caption = q("Question")
    cmbAnswer.Clear
    cmbAnswer.AddItem q("Option1")
    cmbAnswer.AddItem q("Option2")
    cmbAnswer.ListIndex = -1
End Sub

Private Sub btnNext_Click()
    Dim q As Object
    Dim answer As String
    
    If cmbAnswer.ListIndex = -1 Then
        MsgBox "Proszę wybrać odpowiedź.", vbExclamation
        Exit Sub
    End If
    
    answer = cmbAnswer.Value
    Set q = QuestionsDict(CurrentID)
    
    OutputText = OutputText & CurrentID & ": " & answer & vbNewLine
    
    ' Obsługa komentarza
    If Trim(q("Comment")) <> "" Then
        OutputText = OutputText & "INFO: " & q("Comment") & vbNewLine
    End If
    
    ' Zdecyduj o następnym ID na podstawie odpowiedzi
    If answer = q("Option1") Then
        If Left(q("NextIfD"), 2) = "ID" Then
            CurrentID = q("NextIfD")
        Else
            CurrentID = "ID" & q("NextIfD")
        End If
    ElseIf answer = q("Option2") Then
        If Left(q("NextIfE"), 2) = "ID" Then
            CurrentID = q("NextIfE")
        Else
            CurrentID = "ID" & q("NextIfE")
        End If
    Else
        MsgBox "Nieznana odpowiedź.", vbCritical
        Exit Sub
    End If

    ' Obsługa natychmiastowego zakończenia przy "Native/born" (Option1 na ID1)
    If CurrentID = "ID2" And QuestionsDict("ID1")("Option1") = answer Then
        OutputText = OutputText & vbNewLine & "Wynik: All good"
        MsgBox OutputText, vbInformation
        Unload Me
        Exit Sub
    End If

    ' Sprawdź, czy kolejne pytanie istnieje
    If Not QuestionsDict.Exists(CurrentID) Then
        MsgBox "Koniec kwestionariusza." & vbNewLine & vbNewLine & OutputText, vbInformation
        Unload Me
        Exit Sub
    End If
    
    Set q = QuestionsDict(CurrentID)
    
    ' Jeżeli pytanie to "All good" › zakończ
    If Trim(q("Question")) = "All good" Then
        OutputText = OutputText & vbNewLine & "Wynik: All good"
        MsgBox OutputText, vbInformation
        Unload Me
        Exit Sub
    End If

    ShowQuestion CurrentID
End Sub


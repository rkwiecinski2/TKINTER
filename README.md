Option Explicit

Public QuestionsDict As Object
Public CurrentID As String
Public OutputText As String

Sub StartQuestionnaire()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim q As Object
    Dim ID As String
    
    ' Inicjalizacja słownika
    Set QuestionsDict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Wczytaj dane z arkusza do słownika
    For i = 2 To lastRow
        ID = Trim(ws.Cells(i, 1).Value)
        Set q = CreateObject("Scripting.Dictionary")
        
        With q
            .Add "Question", ws.Cells(i, 2).Value
            .Add "Type", ws.Cells(i, 3).Value
            .Add "Option1", ws.Cells(i, 4).Value
            .Add "Option2", ws.Cells(i, 5).Value
            .Add "NextIfD", ws.Cells(i, 6).Value
            .Add "NextIfE", ws.Cells(i, 7).Value
            .Add "Action", ws.Cells(i, 8).Value
            .Add "Comment", ws.Cells(i, 9).Value
        End With
        
        QuestionsDict.Add ID, q
    Next i

    ' Start od ID1
    CurrentID = "ID1"
    OutputText = ""

    ' Pokaż pierwsze pytanie i formularz
    ShowNextQuestion
End Sub

Public Sub ShowNextQuestion()
    Dim q As Object
    
    ' Sprawdź czy istnieje CurrentID
    If Not QuestionsDict.Exists(CurrentID) Then
        MsgBox "Błąd: pytanie o ID '" & CurrentID & "' nie zostało znalezione.", vbCritical
        Exit Sub
    End If
    
    Set q = QuestionsDict(CurrentID)
    
    With QuestionnaireForm
        .lblQuestion.Caption = q("Question")
        .cmbAnswer.Clear
        .cmbAnswer.AddItem q("Option1")
        .cmbAnswer.AddItem q("Option2")
        .cmbAnswer.ListIndex = -1 ' brak zaznaczenia
        .Show
    End With
End Sub





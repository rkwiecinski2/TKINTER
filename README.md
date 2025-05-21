Sub ShowDynamicQuestionnaire()
    Dim wsQ As Worksheet, wsA As Worksheet
    Dim currID As Long, lastRow As Long
    Dim qDict As Object, aRow As Long
    Set wsQ = ThisWorkbook.Sheets("Questions")
    
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("Answers")
    If wsA Is Nothing Then
        Set wsA = ThisWorkbook.Sheets.Add
        wsA.Name = "Answers"
    End If
    wsA.Cells.ClearContents
    wsA.Range("A1:B1").Value = Array("Question", "Answer")
    aRow = 2
    
    Set qDict = CreateObject("Scripting.Dictionary")
    lastRow = wsQ.Cells(wsQ.Rows.Count, "A").End(xlUp).Row
    
    ' Load questions into dictionary
    Dim i As Long
    For i = 2 To lastRow
        qDict(wsQ.Cells(i, 1).Value) = Application.Index(wsQ.Range("A" & i & ":H" & i).Value, 1, 0)
    Next i
    
    currID = 1
    Do While currID <> 0
        Dim qData As Variant
        qData = qDict(currID)
        Dim qType As String: qType = qData(3)
        Dim answer As Variant, nextID As Long
        
        Select Case LCase(qType)
            Case "list"
                answer = Application.InputBox(qData(2) & vbCrLf & "1. " & qData(4) & vbCrLf & "2. " & qData(5), "Select Option (1 or 2)", Type:=1)
                If answer = 1 Then
                    wsA.Cells(aRow, 1).Value = qData(2)
                    wsA.Cells(aRow, 2).Value = qData(4)
                    nextID = qData(6)
                ElseIf answer = 2 Then
                    wsA.Cells(aRow, 1).Value = qData(2)
                    wsA.Cells(aRow, 2).Value = qData(5)
                    nextID = qData(7)
                Else
                    MsgBox "Cancelled or invalid input."
                    Exit Sub
                End If
            
            Case Else
                answer = Application.InputBox(qData(2), "Your Answer")
                If answer = False Then Exit Sub
                wsA.Cells(aRow, 1).Value = qData(2)
                wsA.Cells(aRow, 2).Value = answer
                nextID = 0
        End Select
        
        ' Check if there's an action
        If Trim(qData(8)) <> "" Then
            MsgBox qData(8), vbInformation, "Action"
        End If
        
        aRow = aRow + 1
        currID = nextID
    Loop
    
    MsgBox "All answers recorded!", vbInformation
End Sub

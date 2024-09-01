Sub TransformResponses()
    Dim wsResponses As Worksheet
    Dim wsTransform As Worksheet
    Dim lastRowResponses As Long
    Dim lastRowTransform As Long
    Dim headerRow As Range
    Dim i As Long, j As Long
    Dim questionHeader As String
    Dim colIndex As Long
    Dim answerCount As Integer

    On Error GoTo ErrorHandler

    ' Check if sheets exist and set references
    On Error Resume Next
    Set wsResponses = Sheets("Sheet2") ' Sheet2 contains the responses
    Set wsTransform = Sheets("Sheet6") ' Sheet6 for transformed data
    On Error GoTo ErrorHandler

    ' Verify that wsResponses and wsTransform are correctly set
    If wsResponses Is Nothing Then
        MsgBox "Sheet2 (Responses) not found."
        Exit Sub
    End If

    If wsTransform Is Nothing Then
        ' If Sheet6 does not exist, create it
        Set wsTransform = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTransform.Name = "Sheet6"
    End If

    lastRowResponses = wsResponses.Cells(wsResponses.Rows.Count, "A").End(xlUp).Row
    Set headerRow = wsResponses.Rows(1) ' Header row with questions

    ' Clear old data in Sheet6
    wsTransform.Cells.Clear

    ' Set headers for transformed data in Sheet6
    wsTransform.Cells(1, 1).Value = "Name"
    wsTransform.Cells(1, 2).Value = "Question 1"
    wsTransform.Cells(1, 3).Value = "Answer 1"
    wsTransform.Cells(1, 4).Value = "Question 2"
    wsTransform.Cells(1, 5).Value = "Answer 2"
    wsTransform.Cells(1, 6).Value = "Question 3"
    wsTransform.Cells(1, 7).Value = "Answer 3"
    wsTransform.Cells(1, 8).Value = "Quote"
    wsTransform.Cells(1, 9).Value = "Email"

    ' Transform responses
    lastRowTransform = 2
    For i = 2 To lastRowResponses ' Assuming headers in the first row
        wsTransform.Cells(lastRowTransform, 1).Value = wsResponses.Cells(i, 1).Value ' Name
        
        colIndex = 2
        answerCount = 1
        
        ' Iterate through columns to get questions and answers
        For j = 2 To wsResponses.Columns.Count
            If wsResponses.Cells(i, j).Value <> "" Then
                ' Get the question header from the header row
                questionHeader = headerRow.Cells(1, j).Value
                
                ' Populate question and answer in transformed data
                If answerCount <= 3 Then
                    wsTransform.Cells(lastRowTransform, 2 * answerCount).Value = questionHeader ' Question
                    wsTransform.Cells(lastRowTransform, 2 * answerCount + 1).Value = wsResponses.Cells(i, j).Value ' Answer
                    answerCount = answerCount + 1
                End If
            End If
        Next j
        
        ' Additional fields
        wsTransform.Cells(lastRowTransform, 8).Value = wsResponses.Cells(i, 8).Value ' Quote
        wsTransform.Cells(lastRowTransform, 9).Value = wsResponses.Cells(i, 9).Value ' Email

        lastRowTransform = lastRowTransform + 1
    Next i

    MsgBox "Transformation completed successfully!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub


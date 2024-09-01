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
    Dim email As String
    Dim existingEmail As Range

    On Error GoTo ErrorHandler

    ' Check if sheets exist and set references
    On Error Resume Next
    Set wsResponses = ThisWorkbook.Sheets("Submissions") ' Submissions contains the responses
    Set wsTransform = ThisWorkbook.Sheets("Transformed") ' Transformed for transformed data
    On Error GoTo ErrorHandler

    ' Verify that wsResponses and wsTransform are correctly set
    If wsResponses Is Nothing Then
        MsgBox "Submissions sheet not found."
        Exit Sub
    End If

    If wsTransform Is Nothing Then
        ' If Transformed does not exist, create it
        Set wsTransform = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTransform.Name = "Transformed"
    End If

    lastRowResponses = wsResponses.Cells(wsResponses.Rows.Count, "A").End(xlUp).Row
    Set headerRow = wsResponses.Rows(1) ' Header row with questions

    ' Find the last row in Transformed
    lastRowTransform = wsTransform.Cells(wsTransform.Rows.Count, "A").End(xlUp).Row + 1

    ' Transform responses
    For i = 2 To lastRowResponses ' Assuming headers in the first row
        email = wsResponses.Cells(i, 9).Value ' Email is in column 9

        ' Check if the email already exists in Transformed
        Set existingEmail = wsTransform.Columns(9).Find(What:=email, LookIn:=xlValues, LookAt:=xlWhole)

        If existingEmail Is Nothing Then
            ' If the email doesn't exist, add new data
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
        End If
    Next i

    MsgBox "Transformation completed successfully!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub


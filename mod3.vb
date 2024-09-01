Sub OrganizeDataByHomegroup()
    Dim wsOfficial As Worksheet
    Dim wsData As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowOfficial As Long
    Dim lastRowData As Long
    Dim lastRowTarget As Long
    Dim i As Long, j As Long
    Dim email As String
    Dim homegroup As String
    Dim foundSheet As Boolean
    Dim ws As Worksheet
    Dim cell As Range
    
    On Error GoTo ErrorHandler

    ' Set your sheets
    On Error Resume Next
    Set wsOfficial = Sheets("Sheet5") ' Sheet5 contains the official homegroup data
    Set wsData = Sheets("Sheet6") ' Sheet6 contains the transformed data
    On Error GoTo ErrorHandler

    ' Verify that sheets are correctly set
    If wsOfficial Is Nothing Then
        MsgBox "Sheet5 (Official Data) not found."
        Exit Sub
    End If

    If wsData Is Nothing Then
        MsgBox "Sheet6 (Transformed Data) not found."
        Exit Sub
    End If

    lastRowOfficial = wsOfficial.Cells(wsOfficial.Rows.Count, "A").End(xlUp).Row
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Loop through each email in Sheet6 and find the corresponding homegroup
    For i = 2 To lastRowData ' Assuming headers in the first row
        email = CStr(wsData.Cells(i, 9).Value) ' Email is in column 9
        
        ' Find the corresponding homegroup using email in Sheet5
        homegroup = ""
        For j = 2 To lastRowOfficial
            If CStr(wsOfficial.Cells(j, 5).Value) = email Then ' Email in Sheet5 is in column 5
                homegroup = wsOfficial.Cells(j, 3).Value ' Homegroup in Sheet5 is in column 3
                Exit For
            End If
        Next j
        
        ' If a homegroup is found
        If homegroup <> "" Then
            foundSheet = False
            ' Check if the sheet for the homegroup exists
            For Each ws In ThisWorkbook.Sheets
                If ws.Name = homegroup Then
                    foundSheet = True
                    Set wsTarget = ws
                    Exit For
                End If
            Next ws
            
            ' If the sheet doesn't exist, create it and set the header
            If Not foundSheet Then
                Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsTarget.Name = homegroup
                wsData.Rows(1).Copy Destination:=wsTarget.Rows(1) ' Copy header row from Sheet6
            End If
            
            ' Copy the response data to the appropriate sheet
            lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
            wsTarget.Cells(lastRowTarget, 1).Resize(1, wsData.Columns.Count).Value = wsData.Cells(i, 1).Resize(1, wsData.Columns.Count).Value
            
            ' Logging to help debug
            Debug.Print "Copied data for email " & email & " to sheet " & wsTarget.Name
        Else
            Debug.Print "No homegroup found for email " & email
        End If
    Next i
    
    MsgBox "Data organized by homegroup successfully!"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub


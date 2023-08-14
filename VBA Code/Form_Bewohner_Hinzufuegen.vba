Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    
    ' Get a reference to the database sheet
    Set ws = ThisWorkbook.Worksheets("BewohnerDB")
    
    ' Find the last row in the database
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Size of form
    Me.Height = 195  ' Set the height
    Me.Width = 190   ' Set the width

    ' Clear the combo boxes
    Me.ComboBox_Name.Clear
    Me.ComboBox_Zimmer.Clear
    Me.ComboBox_Ankunft.Clear
    
    ' Populate the name combo box with unique values from the database
    For Each rng In ws.Range("B2:B" & lastRow)
        If Not IsInList(Me.ComboBox_Name, rng.Value) Then
            Me.ComboBox_Name.AddItem rng.Value
        End If
    Next rng
    
    ' Populate the room number combo box with unique values from the database
    For Each rng In ws.Range("D2:D" & lastRow)
        If Not IsInList(Me.ComboBox_Zimmer, rng.Value) Then
            Me.ComboBox_Zimmer.AddItem rng.Value
        End If
    Next rng
    
    ' Populate the start date combo box with unique values from the database
    For Each rng In ws.Range("F2:F" & lastRow)
        If Not IsInList(Me.ComboBox_Ankunft, rng.Value) Then
            Me.ComboBox_Ankunft.AddItem Format(rng.Value, "dd.mm")
        End If
    Next rng
End Sub

Private Sub ComboBox_Name_Change()
    Dim ws As Worksheet
    Dim rng As Range
    Dim matchCount As Long
    Dim matchRow As Long
    
    ' Get a reference to the database sheet
    Set ws = ThisWorkbook.Worksheets("BewohnerDB")

    ' Clear the other combo boxes
    Me.ComboBox_Zimmer.Clear
    Me.ComboBox_Ankunft.Clear

    ' Find matches for the selected name in the database
    Set rng = ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
    matchCount = Application.WorksheetFunction.CountIf(rng, Me.ComboBox_Name.Value)

    ' If there's exactly one match, fill in the other combo boxes
    If matchCount = 1 Then
    matchRow = Application.WorksheetFunction.Match(Me.ComboBox_Name.Value, rng, 0)
    Me.ComboBox_Zimmer.Value = ws.Cells(matchRow + 1, "D").Value
    Me.ComboBox_Ankunft.Value = Format(ws.Cells(matchRow + 1, "F").Value, "dd.mm")
    End If

End Sub

Private Sub ComboBox_Zimmer_Change()
    Dim ws As Worksheet
    Dim rng As Range
    Dim matchCount As Long
    Dim matchRow As Variant
    Dim matchFound As Boolean
    
    ' Get a reference to the database sheet
    Set ws = ThisWorkbook.Worksheets("BewohnerDB")

    ' Clear the start date combo box
    Me.ComboBox_Ankunft.Clear

    ' Find matches for the selected name and room number in the database
    matchFound = False
    For Each rng In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
        If rng.Value = Me.ComboBox_Name.Value And ws.Cells(rng.Row, "D").Value = Me.ComboBox_Zimmer.Value Then
            If matchFound Then
                ' If a second match is found, stop the process and clear the start date combo box
                Me.ComboBox_Ankunft.Clear
                Exit Sub
            Else
                ' If a match is found, remember the row number and set matchFound to True
                matchRow = rng.Row
                matchFound = True
            End If
        End If
    Next rng

    ' If there's exactly one match, fill in the start date combo box
    If matchFound Then
        Me.ComboBox_Ankunft.Value = Format(ws.Cells(matchRow, "F").Value, "dd.mm")
    End If
End Sub

Private Sub CommandButton_OK_Click()

    Dim rng As Range
    
    ' Get a reference to the database sheet
    Set ws = ThisWorkbook.Worksheets("BewohnerDB")

    ' Find the last row in the database
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if the entered name exists in the database
    Dim nameExists As Boolean
    nameExists = False
    For Each rng In ws.Range("B2:B" & lastRow)
        If Me.ComboBox_Name.Value = rng.Value Then
            nameExists = True
            Exit For
        End If
    Next rng
    If Not nameExists Then
        MsgBox "Name existiert nicht.", vbCritical
        Me.ComboBox_Name.SetFocus ' Set the focus on the ComboBox_Name
        Exit Sub
    End If
    
    ' Check if the entered room exists in the database
    Dim roomExists As Boolean
    roomExists = False
    For Each rng In ws.Range("D2:D" & lastRow)
        If CStr(Me.ComboBox_Zimmer.Value) = rng.Value Then
            roomExists = True
            Exit For
        End If
    Next rng
    If Not roomExists Then
        MsgBox "Zimmer existiert nicht.", vbCritical
        Me.ComboBox_Zimmer.SetFocus ' Set the focus on the ComboBox_Zimmer
        Exit Sub
    End If
    
    ' Check if the entered start date exists in the database
    Dim dateExists As Boolean
    dateExists = False
    For Each rng In ws.Range("F2:F" & lastRow)
        If DateValue(Me.ComboBox_Ankunft.Value & "." & Year(Date)) = rng.Value Then
            dateExists = True
            Exit For
        End If
    Next rng
    If Not dateExists Then
        MsgBox "Ankunftsdatum existiert nicht.", vbCritical
        Me.ComboBox_Ankunft.SetFocus ' Set the focus on the ComboBox_Ankunft
        Exit Sub
    End If
    
    ' Check if the date is entered in TextBox_Ende
    If Me.TextBox_Ende.Value = "" Then
        MsgBox "Bitte Enddatum angeben.", vbCritical
        Me.TextBox_Ende.SetFocus ' Set the focus on the TextBox_Ende
        Exit Sub
    Else
        ' Check if the date in TextBox_Ende is in the correct format
        Dim testEndDate As String
        testEndDate = Me.TextBox_Ende.Value & "." & Year(Date)
        If Not IsDate(testEndDate) Then
            MsgBox "Bitte Enddatum im Format dd.mm angeben.", vbCritical
            Me.TextBox_Ende.SetFocus ' Set the focus on the TextBox_Bis
            Exit Sub
        End If
    End If
    
    ' Call the procedure to add the patient when the "OK" button is clicked
    Aufenthalt_Beenden_Routine
    ' Close the UserForm
    Unload Me
End Sub

Private Sub CommandButton_Abbrechen_Click()
    ' Close the UserForm without doing anything when the "Abbrechen" button is clicked
    Unload Me
End Sub

Private Sub Aufenthalt_Beenden_Routine()
    Dim ws As Worksheet
    Dim db As Worksheet
    Dim patientName As String
    Dim roomNumber As String
    Dim startDate As Date
    Dim newEndDate As Date
    Dim dbRow As Long

    ' Get the patient name, room number, start date and new end date
    patientName = Me.ComboBox_Name.Text
    roomNumber = CStr(Me.ComboBox_Zimmer.Value)
    startDate = DateValue(Me.ComboBox_Ankunft.Value & "." & Year(Date))
    newEndDate = DateValue(Me.TextBox_Ende.Value & "." & Year(Date))

    Set db = ThisWorkbook.Worksheets("BewohnerDB")

    ' Find the row in the database for the patient
    For dbRow = 2 To db.Cells(db.Rows.Count, "B").End(xlUp).Row
        If db.Cells(dbRow, "B").Value = patientName And _
           db.Cells(dbRow, "D").Value = roomNumber And _
           CDate(db.Cells(dbRow, "F").Value) = startDate Then
            Exit For
        End If
    Next dbRow

    ' Check if patient was not found
    If dbRow > db.Cells(db.Rows.Count, "B").End(xlUp).Row Then
        MsgBox "Bewohner nicht gefunden.", vbCritical
        Exit Sub
    End If

    ' Get the patients old endDate
    Dim oldEndDate As Date
    oldEndDate = db.Cells(dbRow, 7).Value
    
    ' Check if newEndDate is not between startDate and oldEndDate
    If newEndDate < startDate Or newEndDate > oldEndDate Then
        MsgBox "Neues Enddatum muss zwischen Ankunft und altem Enddatum liegen.", vbCritical
        Exit Sub
    End If

    ' Update the end date in the database
    db.Cells(dbRow, 7).Value = newEndDate

    ' Modify the month sheets
    Dim wsIndex As Integer
    Dim startRow As Long
    Dim startCol As Integer, endCol As Integer

    ' Run through each month from the start month to the old end month
    Debug.Print vbNewLine & "Starting Debug"
    For d = Month(startDate) To Month(oldEndDate)
        wsIndex = d
        Set ws = ThisWorkbook.Worksheets(wsIndex)

        ' Find the row for the room number
        startRow = Application.Match(roomNumber, ws.Columns(2), 0)
        
        ' Find the last column in row 5 of the worksheet
        lastCol = ws.Rows(5).Find(What:="bis", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        lastDayCol = lastCol - 1
        Dim rng As Range
        If Not IsError(startRow) Then
            If d > Month(newEndDate) Then
                ' Determine start and end columns for clearing cells
                startCol = 5
                If d = Month(oldEndDate) Then
                    endCol = day(oldEndDate) + 4 ' 1st day is in the 6th column
                Else
                    endCol = lastDayCol ' Last day of the month
                End If
                
                'Clear and unmerge cells
                Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
                Call DeleteInRange(rng)
                
                ' Clear the patient's name, Pflegegrad and end date
                Dim cellColor As Long
                Dim deletionAllowed As Boolean
                deletionAllowed = True ' Assume that deletion is allowed
                
                For chkCol = endCol + 1 To lastDayCol
                    cellColor = ws.Cells(startRow, chkCol).Interior.Color
                    If cellColor <> 16777215 And cellColor <> -4142 Then
                        deletionAllowed = False ' Set deletionAllowed to False if a colored cell is found
                        Exit For
                    End If
                Next chkCol
                
                If deletionAllowed Then
                    ws.Cells(startRow, 3).ClearContents
                    ws.Cells(startRow, 4).ClearContents
                    ws.Cells(startRow, lastCol).ClearContents
                End If
                
            ElseIf d <= Month(newEndDate) Then
                If d = Month(newEndDate) Then
                    ' Determine start and end columns for clearing cells
                    startCol = day(newEndDate) + 5
                    If d = Month(oldEndDate) Then
                        endCol = day(oldEndDate) + 4 ' 1st day is in the 5th column
                    Else
                        endCol = lastDayCol ' Last day of the month
                    End If
                    
                    'Clear and unmerge cells
                    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
                    Call DeleteInRange(rng)
                    
                    'Remerge Cells before newEndDate
                    If d = Month(startDate) Then
                        startCol = day(startDate) + 4
                    Else
                        startCol = 5
                    End If
                    
                    endCol = day(newEndDate) + 4
                    
                    Debug.Print "Merging cells: Start row " & startRow & ", Start col " & startCol & ", End col " & endCol
                    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
                    rng.Merge
                    'redraw right boarder of merged cell
                    With rng
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlEdgeRight).Weight = xlThin
                    End With
                    
                    'Clear bis column if no other patient
                    deletionAllowed = True
                    For chkCol = day(newEndDate) + 5 To lastDayCol
                        cellColor = ws.Cells(startRow, chkCol).Interior.Color
                        If cellColor <> 16777215 And cellColor <> -4142 Then
                            deletionAllowed = False ' Set deletionAllowed to False if a colored cell is found
                            Exit For
                        End If
                    Next chkCol
                    
                    If deletionAllowed Then
                        ws.Cells(startRow, lastCol).ClearContents
                    End If
                End If
                Debug.Print "Ending operation for month: " & d
        
                ' Write new end date in Bis column for all months from start date to month before newEndDate
                If d < Month(newEndDate) Then
                    ws.Cells(startRow, lastCol).Value = Format(newEndDate, "dd.mm.") ' Write new end date in 'Bis' column
                End If
            End If
        End If
    Next d
End Sub

Function IsInList(cmb As ComboBox, str As String) As Boolean
    Dim I As Long
    For I = 0 To cmb.ListCount - 1
        If cmb.List(I) = str Then
            IsInList = True
            Exit Function
        End If
    Next I
    IsInList = False
End Function

Sub DeleteInRange(deletionRange As Range)
    ' Remember old border styles of first and last cell in the range
    Dim oldTopBorder As XlLineStyle
    Dim oldBottomBorder As XlLineStyle
    Dim oldLeftBorder As XlLineStyle
    Dim oldRightBorder As XlLineStyle

    oldTopBorder = deletionRange.Cells(1, 1).Borders(xlEdgeTop).LineStyle
    oldBottomBorder = deletionRange.Cells(deletionRange.Rows.Count, deletionRange.Columns.Count).Borders(xlEdgeBottom).LineStyle

    oldLeftBorder = deletionRange.Cells(1, 1).Borders(xlEdgeLeft).LineStyle
    oldRightBorder = deletionRange.Cells(deletionRange.Rows.Count, deletionRange.Columns.Count).Borders(xlEdgeRight).LineStyle

    ' Unmerge cells
    deletionRange.MergeCells = False

    ' Clear contents of cell
    deletionRange.ClearContents

    ' Clear colour of cell
    deletionRange.Interior.ColorIndex = 0

    ' Reinstate old top and bottom boarders after unmerging
    With deletionRange
        .Borders(xlEdgeTop).LineStyle = oldTopBorder
        .Borders(xlEdgeBottom).LineStyle = oldBottomBorder
    End With

    ' Set the interior borders to thin for the cells in between
    Dim col As Integer
    For col = 1 To deletionRange.Columns.Count - 1
        With deletionRange.Cells(1, col).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next col
    
    'Reinstate old left and right boarder after unmerging
    deletionRange.Cells(1, 1).Borders(xlEdgeLeft).LineStyle = oldLeftBorder
    deletionRange.Cells(deletionRange.Rows.Count, deletionRange.Columns.Count).Borders(xlEdgeRight).LineStyle = oldRightBorder

End Sub
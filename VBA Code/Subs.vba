Sub TempUnlockRange()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim password As String
    password = "Monika_Schreiber"
    Dim unlockDuration As Date
    unlockDuration = Now + TimeValue("00:02:00")  ' Unlock duration of 2 minutes
    Dim rng As Range
    Dim rng2 As Range
    Dim lastCol As Integer
    Dim found As Range
    
    Dim initialSheet As Worksheet
    Set initialSheet = ActiveSheet  ' Store the currently active sheet

    ' Show the message box
    MsgBox "Manuelle Änderungen bitte nur in Ausnahmefällen nutzen. Sie können die Funktion des Plans beeinflussen und müssen unter Umständen auch wieder manuell rückgangig gemacht werden. Manuelle Änderungen sind jetzt für 2min aktiviert.", vbInformation
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "BewohnerDB" Then
            ' Unprotect the worksheet with the password
            ws.Unprotect password
            
            ' Unlock the ranges
            Set found = ws.Rows(5).Find(What:="bis", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            If Not found Is Nothing Then
                lastCol = found.Column
                Set rng = ws.Range(ws.Cells(6, 3), ws.Cells(40, lastCol))
                rng.Locked = False
            End If
            Set rng2 = ws.Range(ws.Cells(1, 2), ws.Cells(2, 18))
            rng2.Locked = False
            
            ' Protect the worksheet with the password again
            ws.Protect password, UserInterfaceOnly:=True
        End If
    Next ws
    
    initialSheet.Activate  ' Return to the initially active sheet
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    ' Wait until the unlock duration has passed
    Do While Now < unlockDuration
        DoEvents  ' Keeps Excel responsive
    Loop
    
    Set initialSheet = ActiveSheet
    
    ' Show the message box
    MsgBox "Manuelle Änderungen deaktiviert.", vbInformation
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Re-lock the range
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "BewohnerDB" Then
            ' Unprotect the worksheet with the password
            ws.Unprotect password
            
            ' Lock the range
            Set found = ws.Rows(5).Find(What:="bis", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            If Not found Is Nothing Then
                lastCol = found.Column
                Set rng = ws.Range(ws.Cells(6, 3), ws.Cells(40, lastCol))
                rng.Locked = True
            End If
            Set rng2 = ws.Range(ws.Cells(1, 2), ws.Cells(2, 18))
            rng2.Locked = True
            
            ' Protect the worksheet with the password again
            ws.Protect password, UserInterfaceOnly:=True
        End If
    Next ws
    
    initialSheet.Activate
    
ExitSub:
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler: " & Err.Description
    Resume ExitSub
    
End Sub

Sub ClearRooms()
    Dim ws As Worksheet
    Dim lastCol As Integer
    Dim lastRow As Integer
    
    ' Loop through worksheets 1 to 12
    For I = 1 To 12
        Set ws = ThisWorkbook.Worksheets(I)
        
        ' Find the "bis" column and lastRow
        lastCol = ws.Rows(5).Find(What:="bis", LookIn:=xlValues, LookAt:=xlWhole).Column
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        
        ' Delete, clear color, and unboarder all cells below row 5
        With ws.Range(ws.Cells(6, 2), ws.Cells(lastRow, lastCol))
            .ClearContents
            .Interior.Color = xlNone
            .Borders.LineStyle = xlNone
        End With
        
        ' Redraw thick boarder below header
        ws.Range(ws.Cells(5, 2), ws.Cells(5, lastCol)).Borders(xlEdgeBottom).Weight = xlThick
    Next I
End Sub

Sub DeleteButtons()
    Dim ws As Worksheet
    Dim shp As Shape
    For Each ws In ThisWorkbook.Worksheets
        If ws.Index > 1 Then
            For Each shp In ws.Shapes
                If shp.Type = msoFormControl Then
                    If shp.FormControlType = xlButtonControl Then
                        shp.Delete
                    End If
                End If
            Next shp
        End If
    Next ws
End Sub

Sub CopyButtons()
    Dim ws As Worksheet
    Dim sourceWs As Worksheet
    Dim shp As Shape
    Dim newShp As Shape
    Dim assignedMacro As String
    Dim topPos As Double, leftPos As Double

    ' Get the source worksheet (the first one)
    Set sourceWs = ThisWorkbook.Worksheets(1)
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Skip the first worksheet
        If ws.Index > 1 Then
            ' Loop through all shapes in the source worksheet
            For Each shp In sourceWs.Shapes
                ' Check if the shape is a button
                If shp.Type = msoFormControl Then
                    If shp.FormControlType = xlButtonControl Then
                        ' Get the assigned macro
                        assignedMacro = shp.OnAction
                        
                        ' Get button position
                        topPos = shp.Top
                        leftPos = shp.Left
                        
                        ' Copy the button
                        shp.Copy
                        
                        ' Paste the button to the new worksheet at the same position
                        ws.Paste
                        
                        ' Get the new button
                        Set newShp = ws.Shapes(ws.Shapes.Count)
                        
                        ' Assign the same macro to the new button
                        newShp.OnAction = assignedMacro
                        
                        ' Place the button at the same position
                        newShp.Top = topPos
                        newShp.Left = leftPos
                    End If
                End If
            Next shp
        End If
    Next ws
End Sub

Sub PrintButtonNames()
    Dim btn As Object
    For Each btn In ActiveSheet.Buttons
        Debug.Print "Button Name: " & btn.name & ", Button Caption: " & btn.Caption
    Next btn
End Sub

Function EasterDate(Yr As Integer) As Date
    Dim G As Integer
    Dim C As Integer
    Dim H As Integer
    Dim I As Integer
    Dim J As Integer
    Dim L As Integer
    Dim EasterMonth As Integer
    Dim EasterDay As Integer

    G = Yr Mod 19
    C = Yr \ 100
    H = (C - C \ 4 - (8 * C + 13) \ 25 + 19 * G + 15) Mod 30
    I = H - H \ 28 * (1 - 29 \ (H + 1) * (21 - G) \ 11)
    J = (Yr + Yr \ 4 + I + 2 - C + C \ 4) Mod 7
    L = I - J
    EasterMonth = 3 + (L + 40) \ 44
    EasterDay = L + 28 - 31 * (EasterMonth \ 4)

    EasterDate = DateSerial(Yr, EasterMonth, EasterDay)
    Debug.Print "Ostersonntag: " & EasterDate
End Function
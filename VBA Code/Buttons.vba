Sub Bewohner_Hinzufügen_Form()
    bewohner_hinzufügen.Show
End Sub

Sub Bewohner_Löschen_Form()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Get a reference to the database sheet
    Set ws = ThisWorkbook.Worksheets("BewohnerDB")
    
    ' Find the last row in the database
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if the database is empty
    If lastRow <= 1 Then
        MsgBox "Keine Bewohner eingetragen.", vbInformation
        Exit Sub
    End If
    
    Bewohner_Löschen.Show
End Sub

Sub Aufenthalt_Beenden_Form()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Get a reference to the database sheet
    Set ws = ThisWorkbook.Worksheets("BewohnerDB")
    
    ' Find the last row in the database
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if the database is empty
    If lastRow <= 1 Then
        MsgBox "Keine Bewohner eingetragen.", vbInformation
        Exit Sub
    End If
    
    Aufenthalt_Beenden.Show
End Sub

Sub Belegungsplan_Erstellen_Form()
    Belegungsplan_Erstellen.Show
End Sub
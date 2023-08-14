Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ThisWorkbook.Protect Structure:=True, Windows:=False, password:="PW"
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect password:="PW", UserInterfaceOnly:=True
    Next ws
    
    If IsEmpty(Sheet1.Range("B6").Value) Then
        MsgBox "Bei dieser Datei scheint es sich um die Belegungsplan-Schablone zu handeln. Bitte den gewünschten Belegungsplan über die Schaltfläche <Belegungsplan erstellen> erstellen. Die vorgefertigten Stationsnamen und Raumaufteilungen sind vom Stand 2023, können aber manuell angepasst werden.", vbInformation
    End If
End Sub
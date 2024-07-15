'This is the code that is going to run in the background of the to do sheet. It will update the status log and the date last modified. 
' Right now the rows and columns where this information is stored is hardcoding in according to the setup of the original sheet this code was written for however, I am working 
' on a way to make this code more versitile. 


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet, Inser As Worksheet
    Dim rngM As Range, rngAtoQ As Range
    Dim cell As Range
    Dim logCellM As Range, logCellAtoQ As Range
    Dim prevValue As String, newValue As String
   
    'Set referance to the specific worksheet
    Set ws = ThisWorkbook.Sheets("To Do")
    Set Inser = ThisWorkbook.Sheets("Job Insert")
   
    'Check if the changed cells are in column M
    Set rngM = Intersect(Target, ws.Columns("M"))
    If Not rngM Is Nothing Then
        Application.EnableEvents = False 'Disable event handling to prevent infinite loop
        ' Loop through each cell in Column M
        For Each cell In rngM
            newValue = Inser.Cells(cell.Row, "M").Value
           
            If newValue <> cell.Value Then
            'find the corresponding cell in column N to log the changes
            Set logCellM = ws.Cells(cell.Row, "N")
            ' Log the date and changed value
            logCellM.Value = logCellM.Value & vbNewLine & Date & ": " & cell.Value
            End If
           
        Next cell
            Application.EnableEvents = True ' Re-enable event handling
    End If
   
    ' Check is the changes cells are in columns A to Q
    Set rngAtoQ = Intersect(Target, ws.Range("A:Q"))
    If Not rngAtoQ Is Nothing Then
        Application.EnableEvents = False
       
        For Each cell In rngAtoQ
            'Find the corresponding cell in column R to log changes
            Set logCellAtoQ = ws.Cells(cell.Row, "R")
            'Log the date
            logCellAtoQ.Value = Date
        Next cell
       
        Application.EnableEvents = True
    End If
           
            
End Sub

Sub JobInsert()
'This is the job insert code which will allow the user to input new projects and update current ones. See the sheet 1 description in the read me file for more detials as
'this code will be connected to a button on this sheet.

Dim Inser As Worksheet
Dim TD As Worksheet
Dim Dest As Worksheet

'These variables will be counters
Dim i As Long ' Current row in the job insert
Dim j As Long ' Current row in the to do sheet
Dim h As Long ' Current row in the records sheet
Dim k As Long ' Current row in the job set up sheet

Dim Location As Long ' This is the location of any duplicate jobs found (AKA, was it in the to do or records?)
Dim lastRowTD As Long ' This is the first row in the To do sheet without information in it

Dim InserRange As Range ' This is the range we are going to copy in the job insert sheet
Dim TDRange As Range ' This is the range we are going to paste the new job into in the to do sheet

Dim Col As Long ' This is the current Column that we are updateing
Dim FeildType As Long ' This is the feild type information as defined by the user in the setup sheet

Dim currentValue As String ' This is the information currently being stored in any cells that we want to add information to
Dim addValue As String ' This is the information that we are wanting to add on to the old value
Dim newValue As String ' This is going to be the new value we put into the cell

'MsgBox Variables
Dim response As VbMsgBoxResult


'Set the job Insert/search and To do sheet
Set Inser = ThisWorkbook.Sheets("Job Insert")
Set TD = ThisWorkbook.Sheets("To do")
Set Setup = ThisWorkbook.Sheets("SetUp")
Set Records = ThisWorkbook.Sheets("Records")

'Row Counter
 i = 2 'This will track the current row in the job insert sheet

If Inser.Cells(2, 1).Value = "" Then

    MsgBox "Jobs cannot be inserted/updated without an assigned job number to track them by. If you do not know the job number for a job please use the search function below.", vbInformation

End If

'For i = 2 To 9
While Inser.Cells(i, 1).Value <> "" And i < 9

    If Inser.Cells(i, 1).Value = "" Then

        MsgBox "Jobs cannot be inserted/updated without an assigned job number to track them by. If you do not know the job number for a job please use the search function below.", vbInformation
        GoTo Skip
        
    End If
    
'Row Counter
j = 2 ' Current row in the to do sheet
    
'If we do have a job number then we want to search the to do sheet and the records for a duplicate

'Start with the to do sheet, search until there are no more records
While TD.Cells(j, 1).Value <> ""

    If TD.Cells(j, 1).Value = Inser.Cells(i, 1) Then
        Location = 1
        GoTo Question
    End If
    
    'Row Counter
    h = 2 'Current row in the records sheet
    
    'Only if a duplicate was not found in the to do sheet will we also search for it in the records sheet
    While Records.Cells(h, 1).Value <> ""
        
        If Records.Cells(h, 1).Value = Inser.Cells(i, 1) Then
            Location = 2
            GoTo Question
        End If
        
        'Increase row counter
        h = h + 1
    Wend 'End of the while loop searching the records sheet
    
    'Increase row counter
    j = j + 1
    
Wend 'End of while loop searching to do sheet


'If no duplicate job was found in the to do sheet or in the records sheet then we are want to skip to continue and
'enter this job into the next available row in the to do sheet as a new job.
GoTo Continue

    
Question:
'This is the code that runs if a duplicate job is found, the user should be asked if they
'want to continue with this job number or not

'First we want to check the location of the duplicate job so that we can inform the user if the
'job is in progress or closed

If Location = 1 Then

        response = MsgBox("This Job is already on your to do list. Do you want to update it?", vbYesNo + vbQuestion, "Confirmation")
          
          If response = vbYes Then
            'Code to update job in to do
            
           'MsgBox "update if running ", vbInformation
                'Column Counter
                Col = 1 'This is the current Column we are updating information in
                
                'Row counter
                k = 2 'This is the current row in the set up sheet
            
                If Col = 1 Then
                    'Copy paste the job number over every single time
                End If
            
                While Setup.Cells(k, 1).Value <> ""
                    'MsgBox "setup while loop is running ", vbInformation
                    
                    'We need to find the type of the information in the column we are trying to update
                    If Setup.Cells(k, 4).Value = Col Then
                    
                        FeildType = Setup.Cells(k, 2).Value
                        'MsgBox "feild type assigned " & FeildType, vbInformation
                        
                        Select Case FeildType
                        
                        Case Is = 1
                        'Data not to overwrite
                          'MsgBox "case 1", vbInformation
                        
                        If Inser.Cells(i, Col).Value <> "" And TD.Cells(j, Col) = "" Then
                            
                            'Set InserRange = Inser.Range(i & Col & i & Col)
                            'Set TDRange = TD.Range(j & Col & j & Col)
                            
                            'InserRange.Copy TDRange
                            
                            Inser.Cells(i, Col).Copy TD.Cells(j, Col)
                            
                        
                        End If
                        
                        Case Is = 2
                        'Date to overwrite
                          'MsgBox "case 2", vbInformation
                        
                        If Inser.Cells(i, Col).Value <> "" Then
                        
                            'Set InserRange = Inser.Range(i & h & i & h)
                            'Set TDRange = TD.Range(j & h & j & h)
                            
                            'InserRange.Copy TDRange
                            Inser.Cells(i, Col).Copy TD.Cells(j, Col)
                        
                        End If
                        
                        
                        Case Is = 3
                        'Date to add on to
                          'MsgBox "case 3", vbInformation
                        
                        If Inser.Cells(i, Col).Value <> "" Then
                        
                            If TD.Cells(j, Col).Value = "" Then
                            MsgBox "If statment", vbInformation
                            'If the Cell is blank in the to do sheet then we can just copy paste
                            
                                'Set InserRange = Inser.Range(i & h & i & h)
                                'Set TDRange = TD.Range(j & h & j & h)
                            
                                'InserRange.Copy TDRange
                                Inser.Cells(i, Col).Copy TD.Cells(j, Col)
                            
                            Else
                            MsgBox "else statement", vbInformation
                            'If it is not blank then we want to add on information without overwritting the old info.
                            
                            currentValue = TD.Cells(j, Col).Value
                            addValue = Inser.Cells(i, Col).Value
                            newValue = currentValue & vbNewLine & addValue
                            TD.Cells(j, Col).Value = newValue
                        
                            End If
                        
                        End If
                        
                        End Select
                    
                    End If
                    
                    'Increase Counter
                    k = k + 1
                    Col = Col + 1
                
                Wend
            
            
          End If
          
ElseIf Location = 2 Then

    response = MsgBox("This Job is already in your records. You can only update it from the records sheet.", vbYesNo + vbQuestion, "Confirmation")
          If response = vbYes Then
            'Code to update job in records
            GoTo NextRow
            
          End If

End If

'If they did do not want to update the job we need to skip to the next row so that we don't add the duplicate as a new job
GoTo NextRow


Continue:
MsgBox "Continue code", vbInformation
'This is the code for inserting new jobs to the bottom of the to do sheet

'Find the last row of the to do sheet
lastRowTD = TD.Cells(TD.Rows.Count, "A").End(xlUp).Row + 1

'Define the source range (from insert job sheet)
Set InserRange = Inser.Range("A" & i & ":Z" & i)

'Define the destination range (the to do sheet)
Set TDRange = TD.Range("A" & lastRowTD)

'Copy and Paste the information
InserRange.Copy TDRange

    
    
NextRow:
'This is for when a duplicate job is found but the user does not want to update it

'Clear the current row

Inser.Range("A" & i & ":Z" & i).ClearContents

Skip:
'This for when there is an error with a job that is being inserted and we want to skip to the next row but
'not delete the current information in the row so that the error can be fixed

'Increase row counter for the for loop
'Next i
i = i + 1

Wend

End Sub

  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  

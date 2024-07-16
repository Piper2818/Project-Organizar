Sub JobInsert()
'This is the job insert code which will allow the user to input new projects and update current ones. See the sheet 1 description in the read me file for more detials as 
'this code will be connected to a button on this sheet. 

Dim Inser As WorkSheet
Dim TD As WorkSheet 
Dim Dest As Worksheet 

'These variables will be row/column counters 
Dim i As Long 
Dim j As Long 
Dim h As Long 

'These variables will be set according to the location (column) of the information the user needs to track 
Dim A As Long 
Dim B As Long 
Dim C As Long 
Dim D As Long 
Dim E As Long 
Dim F As Long 
Dim G As Long 

'MsgBox Variables 
Dim response As VbMsgBoxResult 

'Set the column variables, this code is currently set up to accommodate excel sheets where the information is organized as follows
'However, I am currently working on a set up code which will allow the user to further personalize their excel sheets depending on use. 
A = 1 ' Date inserted, will not be over written once entered 
B = 2 ' job number 2, will not be over written once entered 
C = 3 ' job number 3, will not be over written once entered 
'Column 4 would be the title/name of the job, in theory
D = 5 ' job status, will only be over written when the user updates the job from the job insert sheet
E = 6 'Flag, is this job flagged as important or not?, will only be over written when the user updates the job from the insert sheet
F = 7 'Notes and/or a summary of the job, will not be overwritten but will be added on to when the user updates the job from the job insert sheet
G = 8 'Date last updated, will but automatically updated (see module 4 code) 
H = 9 'Status log, will be automatically updated (see module 4 code) 

'Set the job Insert/search and To do sheet
Set Inser = ThisWorkbook.Sheets("Job Insert")
Set TD = ThisWorkbook.Sheets("To do")

'Row Counter
 i = 2 'This will track the current row in the job insert sheet 

'If the user has not provided a job type then the system will not know which sheet to sort the job into for records. 
'This if statement can be deleted if the user is only working on one type of job
If Inser.Cells(i, 1).Value <> "" Then 
MsgBox "Jobs Cannot be entered without a job type. This code stops running when it finds the first empty cell in this column, if you have an empty row between jobs you will need to fix this.", vbInformation 
End If 

'This while loop will run from the start to the end of the list given by the user 
While Inser.Cells(i,1).Value <> "" 

  'If the user is trying to enter or update a job without provideing the system with at least one job number to track it by the user should be presented with an error
  'and the code will jump to the next row
  If Inser.Cells(i,A).Value = "" And Inser.Cells(i, B) = "" And Inser.Cells(i,C) = "" 
  MsgBox "Jobs cannot be inserted or updated without at least one job number to identify them by.", vbInformation 
  GoTo skipDuplicate
  End If 

  ElseIf Inser.Cells(i,D) <> "Closed" Then 
    'If the job is not closed then we it is either being entered or updated in the to do sheet. So we will start by searching the to do sheet for a matching job number 
    'but only in the rows that are the correct job type

    Select Case Inser.Cells(i,A).Value
      Case "Job Type 1" 
        'Set the job type sheet 
        Set Dest = ThisWorkbook.Sheets("Job Type 1")
        'Set the row counter 
        j = 2 'j will be the current row in the to do sheet 
        'While loop to search the to do sheet for a match
        While TD.Cells(j,1).Value <> "" 
            If Inser.Cells(i, A).Value <> "" And Inser.Cells(i,A).Value = TD.Cells(j,A).Value Then 
                GoTo Question
            ElseIf Inser.Cells(i, B).Value <> "" And Inser.Cells(i,B) = TD.Cells(i,B).Value Then 
                GoTo Question 
            ElseIf Inser.Cells(i, C).Value <> "" And Inser.Cells(i,C) = TD.Cells(i,C).Value Then 
                GoTo Question 
            'Increase the row counter 
            j = j + 1 
            End If 
      Wend
      GoTo Continue

      Case "Job Type 2" 
        'Set the job type sheet 
        Set Dest = ThisWorkbook.Sheets("Job Type 2")
        'Set the row counter 
        j = 2 'j will be the current row in the to do sheet 
        'While loop to search the to do sheet for a match
        While TD.Cells(j,1).Value <> "" 
            If Inser.Cells(i, A).Value <> "" And Inser.Cells(i,A).Value = TD.Cells(j,A).Value Then 
                GoTo Question
            ElseIf Inser.Cells(i, B).Value <> "" And Inser.Cells(i,B) = TD.Cells(i,B).Value Then 
                GoTo Question 
            ElseIf Inser.Cells(i, C).Value <> "" And Inser.Cells(i,C) = TD.Cells(i,C).Value Then 
                GoTo Question 
            'Increase the row counter 
            j = j + 1 
            End If 
      Wend
      GoTo Continue

      Case "Job Type 3' 
        'Set the job type sheet 
        Set Dest = ThisWorkbook.Sheets("Job Type 3")
        'Set the row counter 
        j = 2 'j will be the current row in the to do sheet 
        'While loop to search the to do sheet for a match
        While TD.Cells(j,1).Value <> "" 
            If Inser.Cells(i, A).Value <> "" And Inser.Cells(i,A).Value = TD.Cells(j,A).Value Then 
                GoTo Question
            ElseIf Inser.Cells(i, B).Value <> "" And Inser.Cells(i,B) = TD.Cells(i,B).Value Then 
                GoTo Question 
            ElseIf Inser.Cells(i, C).Value <> "" And Inser.Cells(i,C) = TD.Cells(i,C).Value Then 
                GoTo Question 
            'Increase the row counter 
            j = j + 1 
            End If 
      Wend
      GoTo Continue
    End Select 
Question: 
          response = MsgBox("This Job Already exsists. Do you want to update it?", vbYesNo + vbQuestion, "Confirmation")
          If response = cbYes Then 
            'Set column counter 
            h = 2 'h will be the current column 
            While h < D ' This while loop will run from Column 2 until the last column holding information which is NOT to be overwritten 

             'First we want to update the columns that should NEVER be overwritten when the user updates the job from the job insert sheet (Ex. Jobs #'S, Title) 
             'Start by checking if this cell is filled in the to do sheet. If it does NOT hold information then we will copy paste the info from the job insert sheet
             If TD.Cells(j,h).Value = "" Then 
              Set InserRange = Inser.Cells(i,h)
              SetTDRange = TD.Cells(j,h)
              InserRange.Copy TDRange
              End If
             h = h + 1 'increase the count to move to next column 
            Wend 

          h = D
          'Now we want to updated the columns that should be overwritten when the user updates the job from the job insert sheet (Ex. status, Flag)
          While h < F
           'Start by checking if this cell is filled in the job insert sheet. If it does hold information then the user is trying to make a change and we need to copy paste
           If Inser.Cells(i,h).Value <> "" Then 
            Set InserRange = Inser.Cells(i,h)
            Set TDRange = Inser.Cells(j,h)
            InserRange.Copy TDRange
           End If 

          h = h + 1 'increase the count to move to next column 
          Wend

          h = F
          'Finally, we want to update the columns that should not be overwritten but need to be added on to (Ex. the notes)  
          While h < G 
           'Start by checking if this cell is filled in the job insert sheet. If it does hold information then the user is trying to add addition information and we want to 
           'put those additions into the to do sheet without losing previous information
            If Inser.Cells(i,h).Value <> "" Then 
             string1 = TD.Cells(j,h).Value 
             string2 = Inser.Cells(i,h).Value
             string1 = string1 & vbNewLine & string2 
             TD.Cells(j,h).Value = string1 
            End If 

          h = h + 1 'increase the count to move to next column 
          Wend 
          
           End If 
         
Continue: 
'Jobs should not be entered without the user choosing a current status for the job


End Sub

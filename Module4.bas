Sub UpdateToDo()

'This code will need to delete any closed jobs, resort the data, update formating, and update Status log

'This code needs to update the correct column in the respective job type sheet, not just add a new row with the closed job so it will first need to find that correct row

 

    Dim TD As Worksheet

    Dim Dest As Worksheet

    Dim i As Long

    Dim J As Long

    Dim h As Long

    Dim lastRow As Long

    Dim lastRowTD As Long

    Dim TDRange As Range

    Dim DestRange As Range

    Dim response As VbMsgBoxResult

   

    'Set the source sheet as the to do sheet because this is where we need to pull the info from

    Set TD = ThisWorkbook.Sheets("To do")

   

    'Set row counter

    i = 2

   

   While TD.Cells(i, 1).Value <> "" ' Start main while

      'MsgBox "while loop row: " & i, vbInformation

   

        If TD.Cells(i, 13).Value = "Closed " Then

          'MsgBox "closed in row: " & i, vbInformation

         

          

          

            Select Case TD.Cells(i, 1).Value ' Start select case

           

' When the job is a damage claim

 

                Case "DMG"

               

                'MsgBox "Damage claim to update in column: " & i, vbInformation

               

                'Move the closed job to the damage claims sheet

                Set Dest = ThisWorkbook.Sheets("Damage Claims")

               

                If TD.Cells(i, 7).Value <> "" Then 'start if 1

                    'Search column 7 in damage claims for a match

                     SearchCol = 7

                ElseIf TD.Cells(i, 8).Value <> "" Then

                    'Search column 8 in damage claims for a match

                    SearchCol = 8

                ElseIf TD.Cells(i, 9).Value <> "" Then

                    'Search column 9 in damage claims for a match

                    SearchCol = 9

                End If ' end if 1

               

                

            If TD.Cells(i, 7).Value = "" Or TD.Cells(i, 8).Value = "" Then ' start if 2

                'Do not let the user close a DC job without a FW# and a DC#

                MsgBox "Damage Claims cannot be closed without a FW# and a DC#", vbInformation

                'Clear the row

                'TD.Range("A" & i & ":Z" & i).ClearContents

                GoTo Continue

               

            End If ' end if 2

   

                'Code to search Damage claims for a match based on the job number found in the last step

                'Will be searching column

       

                J = 2

                While Dest.Cells(J, SearchCol).Value <> "" 'Start of while 2

            

                'When the match is found, copy paste the whole row and jump to the i increament

                'We do not need to search the rest of this column, we now need to continue the search in TD for the next job to update

                    If TD.Cells(i, SearchCol).Value = Dest.Cells(J, SearchCol).Value Then ' start if 3

                        'Set ranges

                        Set TDRange = TD.Range("A" & i & ":T" & i)

                        Set DestRange = Dest.Range("A" & J)

                        'Copy Paste

                        TDRange.Copy DestRange

                        'Delete the information out of to do

                        TD.Rows(i).Delete

                        GoTo Continue

               

                    End If ' end if 3

           

            J = J + 1

        Wend 'end of while 2

               

                

' When the job is an FT3

 

                Case "FT3"

                'Move the closed job to the FT3 sheet

                Set Dest = ThisWorkbook.Sheets("FT3")

               

                

                If TD.Cells(i, 7).Value <> "" Then ' start if 4

                    'Search column 7 in damage claims for a match

                     SearchCol = 7

                ElseIf TD.Cells(i, 11).Value <> "" Then

                    'Search column 11 in damage claims for a match

                    SearchCol = 11

                ElseIf TD.Cells(i, 12).Value <> "" Then

                    'Search column 12 in damage claims for a match

                    SearchCol = 12

                End If ' end if 4

               

                

            If TD.Cells(i, 7).Value = "" Or TD.Cells(i, 11).Value = "" Or TD.Cells(i, 12) = "" Then ' start if 5

                'Do not let the user close a DC job without a FW# and a DC#

                MsgBox "FT# jobs cannot be closed without a FW#, a WFMT#, and a FT3#", vbInformation

                'Clear the row

                'TD.Range("A" & i & ":Z" & i).ClearContents

                GoTo Continue

               

            End If ' start if 5

   

                'Code to search Damage claims for a match based on the job number found in the last step

                'Will be searching column

       

                J = 2

                While Dest.Cells(J, SearchCol).Value <> "" 'Start of while 3

            

                'When the match is found, copy paste the whole row and jump to the i increament

                'We do not need to search the rest of this column, we now need to continue the search in TD for the next job to update

                    If TD.Cells(i, SearchCol).Value = Dest.Cells(J, SearchCol).Value Then ' start if 6

                        'Set ranges

                        Set TDRange = TD.Range("A" & i & ":T" & i)

                        Set DestRange = Dest.Range("A" & J & ":T" & J)

                        'Copy Paste

                        TDRange.Copy DestRange

                        'Delete the information out of to do

                        TD.Rows(i).Delete

                        GoTo Continue

               

                    End If ' end if 6

           

            J = J + 1

        Wend 'end of while 3

               

                

                

        

        

' When the job is a BART Bill

 

                Case "BART"

                'Move the closed job to the BART Bill sheet

                Set Dest = ThisWorkbook.Sheets("BART Bill")

       

                

                If TD.Cells(i, 7).Value <> "" Then ' start if 7

                    'Search column 7 in damage claims for a match

                     SearchCol = 7

                ElseIf TD.Cells(i, 10).Value <> "" Then

                    'Search column 10 in damage claims for a match

                    SearchCol = 10

                ElseIf TD.Cells(i, 11).Value <> "" Then

                    'Search column 11 in damage claims for a match

                    SearchCol = 11

                End If ' end if 7

               

                

            If TD.Cells(i, 7).Value = "" Or TD.Cells(i, 10).Value = "" Or TD.Cells(i, 11) = "" Then ' start if 8

                'Do not let the user close a DC job without a FW# and a DC#

                MsgBox "BART Bill jobs cannot be closed without a FW#, a Armor#, and a WFMT#", vbInformation

                'Clear the row

                'TD.Range("A" & i & ":Z" & i).ClearContents

                GoTo Continue

               

            End If 'end if 8

   

                'Code to search Damage claims for a match based on the job number found in the last step

                'Will be searching column

       

                J = 2

                While Dest.Cells(J, SearchCol).Value <> "" 'Start of while 4

            

                'When the match is found, copy paste the whole row and jump to the i increament

                'We do not need to search the rest of this column, we now need to continue the search in TD for the next job to update

                    If TD.Cells(i, SearchCol).Value = Dest.Cells(J, SearchCol).Value Then ' start if 9

                        'Set ranges

                        Set TDRange = TD.Range("A" & i & ":T" & i)

                        Set DestRange = Dest.Range("A" & J & ":T" & J)

                        'Copy Paste

                        TDRange.Copy DestRange

                        'Delete the information out of to do

                        TD.Rows(i).Delete

                        GoTo Continue

               

                    End If ' end if 9

           

            J = J + 1

        Wend 'end of while 4

               

                

               

        

' When the job is CDFS

               

                Case "CDFS"

                'Move the closed job to the CDFS sheet

                Set Dest = ThisWorkbook.Sheets("CDFS")

               

                

                If TD.Cells(i, 7).Value <> "" Then ' start if 10

                    'Search column 7 in damage claims for a match

                     SearchCol = 7

                ElseIf TD.Cells(i, 10).Value <> "" Then

                    'Search column 10 in damage claims for a match

                    SearchCol = 10

                ElseIf TD.Cells(i, 11).Value <> "" Then

                    'Search column 11 in damage claims for a match

                    SearchCol = 11

                End If ' end if 10

               

                

            If TD.Cells(i, 7).Value = "" Or TD.Cells(i, 10).Value = "" Or TD.Cells(i, 11) = "" Then ' start if 11

                'Do not let the user close a DC job without a FW# and a DC#

                MsgBox "CDFS jobs cannot be closed without a FW#, a Armor#, and WFMT#", vbInformation

                'Clear the row

                'TD.Range("A" & i & ":Z" & i).ClearContents

                GoTo Continue

               

            End If ' end if 11

   

                'Code to search Damage claims for a match based on the job number found in the last step

                'Will be searching column

       

                J = 2

                While Dest.Cells(J, SearchCol).Value <> "" 'Start of while 5

            

                'When the match is found, copy paste the whole row and jump to the i increament

                'We do not need to search the rest of this column, we now need to continue the search in TD for the next job to update

                    If TD.Cells(i, SearchCol).Value = Dest.Cells(J, SearchCol).Value Then ' start if 12

                        'Set ranges

                        Set TDRange = TD.Range("A" & i & ":T" & i)

                        Set DestRange = Dest.Range("A" & J & ":T" & J)

                        'Copy Paste

                        TDRange.Copy DestRange

                        'Delete the information out of to do

                        TD.Rows(i).Delete

                        GoTo Continue

               

                    End If ' end if 12

           

            J = J + 1

        Wend 'end of while 5

               

            End Select ' end select

       

            

        Else

Continue:

        'Do not update the row count unless the job in the row was not closed because when the closed job is deleted from the to do list all the data is going to shift

        i = i + 1

       

        End If ' end main if

       

    Wend ' Start main while

       

 

 

 

End Sub

 

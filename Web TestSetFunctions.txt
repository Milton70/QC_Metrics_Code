Attribute VB_Name = "TestSetFunctions"
Public Function GetTestsByFunctionalArea()
Dim iCount As Integer
Dim iTotCount As Integer
iCount = 0
iTotCount = 0
    
    ' Count the critical defects by test phase and add them to the data sheet.
    arrTestPhaseKeys = CountTestsByFunctionalArea()
    
    If UBound(arrTestPhaseKeys) <= 0 Then
        Exit Function
    End If
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "TestFunctionalAreaTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestFunctionalArea.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "PassedPoints") > 0 Then
            '   Loop round our array
            For i = 0 To UBound(arrTestPhaseKeys)
                '   Split the file
                mySplit = Split(arrTestPhaseKeys(i), "|")
                If mySplit(0) = "Passed" Then
                    '   Write the values
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                End If
            Next
        Else
            If InStr(1, rc, "FailedPoints") > 0 Then
                '   Loop round our array
                For i = 0 To UBound(arrTestPhaseKeys)
                    '   Split the file
                    mySplit = Split(arrTestPhaseKeys(i), "|")
                    If mySplit(0) = "Failed" Then
                        '   Write the values
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                    End If
                Next
            Else
                If InStr(1, rc, "NotCompletedPoints") > 0 Then
                    '   Loop round our array
                    For i = 0 To UBound(arrTestPhaseKeys)
                        '   Split the file
                        mySplit = Split(arrTestPhaseKeys(i), "|")
                        If mySplit(0) = "Not Completed" Then
                            '   Write the values
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                        End If
                    Next
                Else
                    If InStr(1, rc, "NoRunPoints") > 0 Then
                        '   Loop round our array
                        For i = 0 To UBound(arrTestPhaseKeys)
                            '   Split the file
                            mySplit = Split(arrTestPhaseKeys(i), "|")
                            If mySplit(0) = "No Run" Then
                                '   Write the values
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                            End If
                        Next
                    Else
                        myDest.WriteLine rc
                    End If
                End If
            End If
        End If
    Loop
    
    '   Close the files
    myDest.Close
    mySource.Close
    
    '   See how many functional areas we've got in the dictionary and this will determine the size of the graph
    myTemp = ""
    For Each Ele In arrTestPhaseKeys
        mySplit = Split(Ele, "|")
        If mySplit(0) = "Passed" Then
            If myTemp = "" Or myTemp <> mySplit(1) Then
                myTemp = mySplit(1)
                iCount = iCount + 1
            End If
        End If
    Next
    myTemp = ""
    For Each Ele In arrTestPhaseKeys
        iTotCount = iTotCount + objTestFunctionalAreaDictionary(Ele)
    Next
    
    '   Open the file and change the width value
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestFunctionalArea.aspx", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    '   Depending on the number, change the value
    Select Case iCount
        Case 1, 2, 3
            strText = Replace(strText, "SetWidth", Chr(34) & "75" & Chr(34))
        Case 4, 5, 6
            strText = Replace(strText, "SetWidth", Chr(34) & "80" & Chr(34))
        Case Else
            strText = Replace(strText, "SetWidth", Chr(34) & "100" & Chr(34))
    End Select
    Select Case iTotCount
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            strText = Replace(strText, "ReplaceYAxis", "Interval=" & Chr(34) & "1" & Chr(34))
        Case Else
            strText = Replace(strText, "ReplaceYAxis", "")
    End Select
    
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestFunctionalArea.aspx", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Create the supporting test set page and include the graph just created
    Set myFile = fso.OpenTextFile(strTemplatePath & "TestSetDetailsTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    '   Replace the graph
    strText = Replace(strText, "TestsByFunctionalArea", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestFunctionalArea.aspx" & Chr(34))
    strText = Replace(strText, "strHeader", "Supporting Test Script Details for " & strHeader)
    
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Copy the hyperlink table file to the folder
    fso.CopyFile strTemplatePath & "TestFunctionalAreaTableTemplate.txt", strFolderPath & strPathandFileName & "-TestFunctionalAreaTable.asp"
    '   Now do the hyperlink table
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestFunctionalAreaTable.asp", ForAppending)
    myTemp = ""
    For i = 0 To UBound(arrTestPhaseKeys)
        '   Split the file
        mySplit = Split(arrTestPhaseKeys(i), "|")
        If mySplit(0) = "Failed" Then
            If myTemp = "" Or myTemp <> mySplit(1) Then
                myTemp = mySplit(1)
                myFile.WriteLine "<tr><td>" & mySplit(1) & "</td><td>" & objTestFunctionalAreaDictionary("Passed|" & mySplit(1)) & "</td><td>" & objTestFunctionalAreaDictionary(arrTestPhaseKeys(i)) & "</td><td>" & objTestFunctionalAreaDictionary("Not Completed|" & mySplit(1)) & "</td><td>" & objTestFunctionalAreaDictionary("No Run|" & mySplit(1)) & "</td></tr>"
            End If
        Else
            Exit For
        End If
    Next
    '   Finish off html
    myFile.WriteLine "</table></table></body></html>"
    myFile.Close

End Function
Public Function GetTestSets()
Dim intPlannedCount As Integer, intActualCount As Integer, intPassedCount As Integer, intFailedCount As Integer, intNCCount As Integer, intReRun As Integer
Dim arrStatus As Variant

    '   Set up the new dictionary
    Set objTestSetDictionary = New Dictionary
    
    '   Get Total number of tests planned
    intPlannedCount = 0
    intPlannedCount = CountTestsRunByDate("Planned-Baseline", "")
    objTestSetDictionary.Add "Total No Scripts", intPlannedCount
    
    '   Set the total planned global variable
    intTotalPlannedTests = intPlannedCount
    
    If intTotalPlannedTests = 0 Then
        '   See if test set status file exists
        If fso.FileExists(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt") = True Then
            '   Delete it
            fso.DeleteFile strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", True
        End If
        '   Create it again, from the error template
        fso.CopyFile strTemplatePath & "TestSetsMissingTemplate.txt", strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt"
    
        '   Open the file to read the data
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        '   Change the header
        strText = Replace(strText, "strHeader", "No Test Script data for " & strHeader)
    
        '   Write the data outr
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", ForWriting, True)
        myFile.WriteLine strText
        myFile.Close
    
        '   Set the don't run flag
        blnDontRun = True
        blnNoTestSets = True
        Exit Function
    End If

    '   Set blnNoTestSets to false if we're here
    blnNoTestSets = False
    
    '   Get the Total number of tests passed
    intPassedCount = 0
    intPassedCount = CountTestsRunByDate("Status", "", "Passed")
    objTestSetDictionary.Add "Total Passed", intPassedCount
    
    '   Get the Total number of tests failed
    intFailedCount = 0
    intFailedCount = CountTestsRunByDate("Status", "", "Failed")
    objTestSetDictionary.Add "Total Failed", intFailedCount
    
    '   Get the Total number of tests Not Completed
    intNCCount = 0
    intNCCount = CountTestsRunByDate("Status", "", "Not Completed")
    objTestSetDictionary.Add "Total Not Completed", intNCCount
    
    '   Get the Total number of tests Not Run
    intNRCount = 0
    intNRCount = CountTestsRunByDate("Status", "", "No Run")
    objTestSetDictionary.Add "Total No Run", intNRCount
    
    '  Calculate the executed value
    intExecuted = 0
    intExecuted = intPassedCount + intFailedCount
    objTestSetDictionary.Add "Total Executed", intExecuted
    
    '   Now calculate the percentage values
    If intTotalPlannedTests = 0 Then
        iNoRunPercent = "0"
        iExecutedPercent = "0"
        iPassedPercent = "0"
        iFailedPercent = "0"
        iNCPercent = "0"
        iTotalPercent = "0"
        iExecutedPassed = "0"
        iExecutedFailed = "0"
        iExecutedNC = "0"
    Else
        iNoRunPercent = Round(intNRCount / intTotalPlannedTests * 100, 2)
        iExecutedPercent = Round(intExecuted / intTotalPlannedTests * 100, 2)
        iPassedPercent = Round(intPassedCount / intTotalPlannedTests * 100, 2)
        iFailedPercent = Round(intFailedCount / intTotalPlannedTests * 100, 2)
        iNCPercent = Round(intNCCount / intTotalPlannedTests * 100, 2)
        iTotalPercent = iNoRunPercent + iExecutedPercent
        If intPassedCount > 0 And intExecuted > 0 Then
            iExecutedPassed = Round(intPassedCount / intExecuted * 100, 2)
        Else
            iExecutedPassed = "0"
        End If
        If intFailedCount > 0 And intExecuted > 0 Then
            iExecutedFailed = Round(intFailedCount / intExecuted * 100, 2)
        Else
            iExecutedFailed = "0"
        End If
        If intNCCount > 0 And intExecuted > 0 Then
            iExecutedNC = Round(intNCCount / intExecuted * 100, 2)
        Else
            iExecutedNC = "0"
        End If
    End If
    dblGlobalExecutedPercent = iExecutedPercent
    dblGlobalPlannedPassedPercent = iPassedPercent
    dblGlobalPlannedFailedPercent = iFailedPercent
    dblGlobalExecutedPassedPercent = iExecutedPassed
    dblGlobalExecutedFailedPercent = iExecutedFailed
   
    '   Get the Priority details
    intPlannedP1 = CountTestsRunByDate("Planned-Priority", "", "", "1-Urgent")
    intPlannedP2 = CountTestsRunByDate("Planned-Priority", "", "", "2-High")
    intPlannedP3 = CountTestsRunByDate("Planned-Priority", "", "", "3-Medium")
    intPlannedP4 = CountTestsRunByDate("Planned-Priority", "", "", "4-Low")
    objTestSetDictionary.Add "Planned Priority 1", intPlannedP1
    objTestSetDictionary.Add "Planned Priority 2", intPlannedP2
    objTestSetDictionary.Add "Planned Priority 3", intPlannedP3
    objTestSetDictionary.Add "Planned Priority 4", intPlannedP4

    intPassedP1 = CountTestsRunByDate("Status-Priority", "", "Passed", "1-Urgent")
    intPassedP2 = CountTestsRunByDate("Status-Priority", "", "Passed", "2-High")
    intPassedP3 = CountTestsRunByDate("Status-Priority", "", "Passed", "3-Medium")
    intPassedP4 = CountTestsRunByDate("Status-Priority", "", "Passed", "4-Low")
    objTestSetDictionary.Add "Passed Priority 1", intPassedP1
    objTestSetDictionary.Add "Passed Priority 2", intPassedP2
    objTestSetDictionary.Add "Passed Priority 3", intPassedP3
    objTestSetDictionary.Add "Passed Priority 4", intPassedP4

    intFailedP1 = CountTestsRunByDate("Status-Priority", "", "Failed", "1-Urgent")
    intFailedP2 = CountTestsRunByDate("Status-Priority", "", "Failed", "2-High")
    intFailedP3 = CountTestsRunByDate("Status-Priority", "", "Failed", "3-Medium")
    intFailedP4 = CountTestsRunByDate("Status-Priority", "", "Failed", "4-Low")
    objTestSetDictionary.Add "Failed Priority 1", intFailedP1
    objTestSetDictionary.Add "Failed Priority 2", intFailedP2
    objTestSetDictionary.Add "Failed Priority 3", intFailedP3
    objTestSetDictionary.Add "Failed Priority 4", intFailedP4
    
    intNCP1 = CountTestsRunByDate("Status-Priority", "", "Not Completed", "1-Urgent")
    intNCP2 = CountTestsRunByDate("Status-Priority", "", "Not Completed", "2-High")
    intNCP3 = CountTestsRunByDate("Status-Priority", "", "Not Completed", "3-Medium")
    intNCP4 = CountTestsRunByDate("Status-Priority", "", "Not Completed", "4-Low")
    objTestSetDictionary.Add "Not Completed Priority 1", intNCP1
    objTestSetDictionary.Add "Not Completed Priority 2", intNCP2
    objTestSetDictionary.Add "Not Completed Priority 3", intNCP3
    objTestSetDictionary.Add "Not Completed Priority 4", intNCP4
    
    intNRP1 = CountTestsRunByDate("Status-Priority", "", "No Run", "1-Urgent")
    intNRP2 = CountTestsRunByDate("Status-Priority", "", "No Run", "2-High")
    intNRP3 = CountTestsRunByDate("Status-Priority", "", "No Run", "3-Medium")
    intNRP4 = CountTestsRunByDate("Status-Priority", "", "No Run", "4-Low")
    objTestSetDictionary.Add "No Run Priority 1", intNRP1
    objTestSetDictionary.Add "No Run Priority 2", intNRP2
    objTestSetDictionary.Add "No Run Priority 3", intNRP3
    objTestSetDictionary.Add "No Run Priority 4", intNRP4
    
    '   Set up the executed priorities
    objTestSetDictionary.Add "Executed Priority 1", intPassedP1 + intFailedP1
    objTestSetDictionary.Add "Executed Priority 2", intPassedP2 + intFailedP2
    objTestSetDictionary.Add "Executed Priority 3", intPassedP3 + intFailedP3
    objTestSetDictionary.Add "Executed Priority 4", intPassedP4 + intFailedP4
    
    '   See if we've got no total planned tests
    If intTotalPlannedTests = 0 Then
        iP1Total = "0"
        iP2Total = "0"
        iP3Total = "0"
        iP4Total = "0"
    Else
        '   Calculate the priority percentages
        iP1Total = Round(intPlannedP1 / intTotalPlannedTests * 100, 2)
        iP2Total = Round(intPlannedP2 / intTotalPlannedTests * 100, 2)
        iP3Total = Round(intPlannedP3 / intTotalPlannedTests * 100, 2)
        iP4Total = Round(intPlannedP4 / intTotalPlannedTests * 100, 2)
    End If
    '   See if we've got no priority planned
    If intPlannedP1 = 0 Then
        iP1NoRun = "0"
    Else
        iP1NoRun = Round(intNRP1 / intPlannedP1 * 100, 2)
    End If
    If intPlannedP2 = 0 Then
        iP2NoRun = "0"
    Else
        iP2NoRun = Round(intNRP2 / intPlannedP2 * 100, 2)
    End If
    If intPlannedP3 = 0 Then
        iP3NoRun = "0"
    Else
        iP3NoRun = Round(intNRP3 / intPlannedP3 * 100, 2)
    End If
    If intPlannedP4 = 0 Then
        iP4NoRun = "0"
    Else
        iP4NoRun = Round(intNRP4 / intPlannedP4 * 100, 2)
    End If
    If intPlannedP1 = 0 Then
        iP1Executed = "0"
    Else
        iP1Executed = Round((intPassedP1 + intFailedP1) / intPlannedP1 * 100, 2)
    End If
    If intPlannedP2 = 0 Then
        iP2Executed = "0"
    Else
        iP2Executed = Round((intPassedP2 + intFailedP2) / intPlannedP2 * 100, 2)
    End If
    If intPlannedP3 = 0 Then
        iP3Executed = "0"
    Else
        iP3Executed = Round((intPassedP3 + intFailedP3) / intPlannedP3 * 100, 2)
    End If
    If intPlannedP4 = 0 Then
        iP4Executed = "0"
    Else
        iP4Executed = Round((intPassedP4 + intFailedP4) / intPlannedP4 * 100, 2)
    End If
    If (intPassedP1 + intFailedP1) = 0 Then
        iP1Passed = "0"
    Else
        iP1Passed = Round(intPassedP1 / (intPassedP1 + intFailedP1) * 100, 2)
    End If
    If (intPassedP2 + intFailedP2) = 0 Then
        iP2Passed = "0"
    Else
        iP2Passed = Round(intPassedP2 / (intPassedP2 + intFailedP2) * 100, 2)
    End If
    If (intPassedP3 + intFailedP3) = 0 Then
        iP3Passed = "0"
    Else
        iP3Passed = Round(intPassedP3 / (intPassedP3 + intFailedP3) * 100, 2)
    End If
    If (intPassedP4 + intFailedP4) = 0 Then
        iP4Passed = "0"
    Else
        iP4Passed = Round(intPassedP4 / (intPassedP4 + intFailedP4) * 100, 2)
    End If
    If (intPassedP1 + intFailedP1) = 0 Then
        iP1Failed = "0"
    Else
        iP1Failed = Round(intFailedP1 / (intPassedP1 + intFailedP1) * 100, 2)
    End If
    If (intPassedP2 + intFailedP2) = 0 Then
        iP2Failed = "0"
    Else
        iP2Failed = Round(intFailedP2 / (intPassedP2 + intFailedP2) * 100, 2)
    End If
    If (intPassedP3 + intFailedP3) = 0 Then
        iP3Failed = "0"
    Else
        iP3Failed = Round(intFailedP3 / (intPassedP3 + intFailedP3) * 100, 2)
    End If
    If (intPassedP4 + intFailedP4) = 0 Then
        iP4Failed = "0"
    Else
        iP4Failed = Round(intFailedP4 / (intPassedP4 + intFailedP4) * 100, 2)
    End If
    If (intPassedP1 + intFailedP1) = 0 Then
        iP1NC = "0"
    Else
        iP1NC = Round(intNCP1 / (intPassedP1 + intFailedP1) * 100, 2)
    End If
    If (intPassedP2 + intFailedP2) = 0 Then
        iP2NC = "0"
    Else
        iP2NC = Round(intNCP2 / (intPassedP2 + intFailedP2) * 100, 2)
    End If
    If (intPassedP3 + intFailedP3) = 0 Then
        iP3NC = "0"
    Else
        iP3NC = Round(intNCP3 / (intPassedP3 + intFailedP3) * 100, 2)
    End If
    If (intPassedP4 + intFailedP4) = 0 Then
        iP4NC = "0"
    Else
        iP4NC = Round(intNCP4 / (intPassedP4 + intFailedP4) * 100, 2)
    End If
    
    '   Get the Total number of planned tests Today
    intPlannedCount = 0
    intPlannedCount = CountTestsRunByDate("Replanned", TodaysDate)
    objTestSetDictionary.Add "Planned Today", intPlannedCount
        
    '   Get the Total number of tests passed Today
    intPassedCount = 0
    intPassedCount = CountTestsRunByDate("Status", TodaysDate, "Passed")
    objTestSetDictionary.Add "Passed Today", intPassedCount

    '   Get the Total number of tests failed Today
    intFailedCount = 0
    intFailedCount = CountTestsRunByDate("Status", TodaysDate, "Failed")
    objTestSetDictionary.Add "Failed Today", intFailedCount

    '   Get the Total number of tests Not Completed Today
    intNCCount = 0
    intNCCount = CountTestsRunByDate("Status", TodaysDate, "Not Completed")
    objTestSetDictionary.Add "Not Completed Today", intNCCount

    '   Get the Total number of tests Not Run Today
    intNRCount = 0
    intNRCount = CountTestsRunByDate("Status", TodaysDate, "No Run Daily")
    objTestSetDictionary.Add "No Run Today", intNRCount
    
    '   Set up the Executed value
    objTestSetDictionary.Add "Executed Today", intFailedCount + intPassedCount
    
    '   Output the data
    strWeekDay = GetWeekday(TodaysDate)
    
    '   Get this week's data
    iWeekDay = Weekday(TodaysDate)
    Select Case iWeekDay
        Case 2
            dateThisWeek = TodaysDate
        Case 3
            dateThisWeek = DateAdd("d", -1, TodaysDate)
        Case 4
            dateThisWeek = DateAdd("d", -2, TodaysDate)
        Case 5
            dateThisWeek = DateAdd("d", -3, TodaysDate)
        Case 6
            dateThisWeek = DateAdd("d", -4, TodaysDate)
    End Select
    
    '   Get the start weekday
    strStartWeekday = GetWeekday(dateThisWeek)
    
    '   Get the Total number of test planned this week
    intPlannedCount = 0
    intPlannedCount = CountTestsRunByDate("Replanned", ">= " & dateThisWeek & " And <=" & TodaysDate)
    objTestSetDictionary.Add "Planned Week", intPlannedCount
    
     '   Get the Total number of tests passed this week
    intPassedCount = 0
    intPassedCount = CountTestsRunByDate("Status", ">= " & dateThisWeek & " And <=" & TodaysDate, "Passed")
    objTestSetDictionary.Add "Passed Week", intPassedCount
    
    '   Get the Total number of tests failed this week
    intFailedCount = 0
    intFailedCount = CountTestsRunByDate("Status", ">= " & dateThisWeek & " And <=" & TodaysDate, "Failed")
    objTestSetDictionary.Add "Failed Week", intFailedCount
    
    '   Get the Total number of tests Not Completed this week
    intNCCount = 0
    intNCCount = CountTestsRunByDate("Status", ">= " & dateThisWeek & " And <=" & TodaysDate, "Not Completed")
    objTestSetDictionary.Add "Not Completed Week", intNCCount
    
    '   Get the Total number of tests Not Run this week
    intNRCount = 0
    intNRCount = CountTestsRunByDate("Status", ">= " & dateThisWeek & " And <=" & TodaysDate, "No Run Weekly")
    objTestSetDictionary.Add "No Run Week", intNRCount
    
    '   Calculate executed week
    objTestSetDictionary.Add "Executed Week", intFailedCount + intPassedCount
    
    '  Open the template file
    Set myFile = fso.OpenTextFile(strTemplatePath & "TestSetStatusTemplate.txt", ForReading)
    '   Get the text
    strText = myFile.ReadAll
    '   Close this file
    myFile.Close
    
    '   Start Replacing the text in the file
    strNewText = Replace(strText, "strHeader", strHeader)
    strNewText = Replace(strNewText, "TotalNoScripts", objTestSetDictionary.Item("Total No Scripts"))
    strNewText = Replace(strNewText, "iTotalPercent", iTotalPercent)
    strNewText = Replace(strNewText, "TotalNoRun", objTestSetDictionary.Item("Total No Run"))
    strNewText = Replace(strNewText, "iNoRunPercent", iNoRunPercent)
    strNewText = Replace(strNewText, "TotalExecuted", objTestSetDictionary.Item("Total Executed"))
    strNewText = Replace(strNewText, "iExecutedPercent", iExecutedPercent)
    strNewText = Replace(strNewText, "TotalPassed", objTestSetDictionary.Item("Total Passed"))
    strNewText = Replace(strNewText, "iPassedPercent", iPassedPercent)
    strNewText = Replace(strNewText, "iExecutedPassed", iExecutedPassed)
    strNewText = Replace(strNewText, "TotalFailed", objTestSetDictionary.Item("Total Failed"))
    strNewText = Replace(strNewText, "iFailedPercent", iFailedPercent)
    strNewText = Replace(strNewText, "iExecutedFailed", iExecutedFailed)
    strNewText = Replace(strNewText, "TotalNotCompleted", objTestSetDictionary.Item("Total Not Completed"))
    strNewText = Replace(strNewText, "iNCPercent", iNCPercent)
    strNewText = Replace(strNewText, "iExecutedNC", iExecutedNC)
    strNewText = Replace(strNewText, "strWeekday", strWeekDay)
    strNewText = Replace(strNewText, "TodaysDate", TodaysDate)
    strNewText = Replace(strNewText, "PlannedToday", objTestSetDictionary.Item("Planned Today"))
    strNewText = Replace(strNewText, "NoRunToday", objTestSetDictionary.Item("No Run Today"))
    strNewText = Replace(strNewText, "ExecutedToday", objTestSetDictionary.Item("Executed Today"))
    strNewText = Replace(strNewText, "PassedToday", objTestSetDictionary.Item("Passed Today"))
    strNewText = Replace(strNewText, "FailedToday", objTestSetDictionary.Item("Failed Today"))
    strNewText = Replace(strNewText, "NotCompletedToday", objTestSetDictionary.Item("Not Completed Today"))
    strNewText = Replace(strNewText, "strStartWeekday", strStartWeekday)
    strNewText = Replace(strNewText, "dateThisWeek", dateThisWeek)
    strNewText = Replace(strNewText, "PlannedWeek", objTestSetDictionary.Item("Planned Week"))
    strNewText = Replace(strNewText, "NoRunWeek", objTestSetDictionary.Item("No Run Week"))
    strNewText = Replace(strNewText, "ExecutedWeek", objTestSetDictionary.Item("Executed Week"))
    strNewText = Replace(strNewText, "PassedWeek", objTestSetDictionary.Item("Passed Week"))
    strNewText = Replace(strNewText, "FailedWeek", objTestSetDictionary.Item("Failed Week"))
    strNewText = Replace(strNewText, "NotCompletedWeek", objTestSetDictionary.Item("Not Completed Week"))
    strNewText = Replace(strNewText, "PlannedPriority1", objTestSetDictionary.Item("Planned Priority 1"))
    strNewText = Replace(strNewText, "PlannedPriority2", objTestSetDictionary.Item("Planned Priority 2"))
    strNewText = Replace(strNewText, "PlannedPriority3", objTestSetDictionary.Item("Planned Priority 3"))
    strNewText = Replace(strNewText, "PlannedPriority4", objTestSetDictionary.Item("Planned Priority 4"))
    strNewText = Replace(strNewText, "iP1Total", iP1Total)
    strNewText = Replace(strNewText, "iP2Total", iP2Total)
    strNewText = Replace(strNewText, "iP3Total", iP3Total)
    strNewText = Replace(strNewText, "iP4Total", iP4Total)
    strNewText = Replace(strNewText, "NoRunPriority1", objTestSetDictionary.Item("No Run Priority 1"))
    strNewText = Replace(strNewText, "NoRunPriority2", objTestSetDictionary.Item("No Run Priority 2"))
    strNewText = Replace(strNewText, "NoRunPriority3", objTestSetDictionary.Item("No Run Priority 3"))
    strNewText = Replace(strNewText, "NoRunPriority4", objTestSetDictionary.Item("No Run Priority 4"))
    strNewText = Replace(strNewText, "iP1NoRun", iP1NoRun)
    strNewText = Replace(strNewText, "iP2NoRun", iP2NoRun)
    strNewText = Replace(strNewText, "iP3NoRun", iP3NoRun)
    strNewText = Replace(strNewText, "iP4NoRun", iP4NoRun)
    strNewText = Replace(strNewText, "ExecutedPriority1", objTestSetDictionary.Item("Executed Priority 1"))
    strNewText = Replace(strNewText, "ExecutedPriority2", objTestSetDictionary.Item("Executed Priority 2"))
    strNewText = Replace(strNewText, "ExecutedPriority3", objTestSetDictionary.Item("Executed Priority 3"))
    strNewText = Replace(strNewText, "ExecutedPriority4", objTestSetDictionary.Item("Executed Priority 4"))
    strNewText = Replace(strNewText, "iP1Executed", iP1Executed)
    strNewText = Replace(strNewText, "iP2Executed", iP2Executed)
    strNewText = Replace(strNewText, "iP3Executed", iP3Executed)
    strNewText = Replace(strNewText, "iP4Executed", iP4Executed)
    strNewText = Replace(strNewText, "PassedPriority1", objTestSetDictionary.Item("Passed Priority 1"))
    strNewText = Replace(strNewText, "PassedPriority2", objTestSetDictionary.Item("Passed Priority 2"))
    strNewText = Replace(strNewText, "PassedPriority3", objTestSetDictionary.Item("Passed Priority 3"))
    strNewText = Replace(strNewText, "PassedPriority4", objTestSetDictionary.Item("Passed Priority 4"))
    strNewText = Replace(strNewText, "iP1Passed", iP1Passed)
    strNewText = Replace(strNewText, "iP2Passed", iP2Passed)
    strNewText = Replace(strNewText, "iP3Passed", iP3Passed)
    strNewText = Replace(strNewText, "iP4Passed", iP4Passed)
    strNewText = Replace(strNewText, "FailedPriority1", objTestSetDictionary.Item("Failed Priority 1"))
    strNewText = Replace(strNewText, "FailedPriority2", objTestSetDictionary.Item("Failed Priority 2"))
    strNewText = Replace(strNewText, "FailedPriority3", objTestSetDictionary.Item("Failed Priority 3"))
    strNewText = Replace(strNewText, "FailedPriority4", objTestSetDictionary.Item("Failed Priority 4"))
    strNewText = Replace(strNewText, "iP1Failed", iP1Failed)
    strNewText = Replace(strNewText, "iP2Failed", iP2Failed)
    strNewText = Replace(strNewText, "iP3Failed", iP3Failed)
    strNewText = Replace(strNewText, "iP4Failed", iP4Failed)
    strNewText = Replace(strNewText, "NotCompletedPriority1", objTestSetDictionary.Item("Not Completed Priority 1"))
    strNewText = Replace(strNewText, "NotCompletedPriority2", objTestSetDictionary.Item("Not Completed Priority 2"))
    strNewText = Replace(strNewText, "NotCompletedPriority3", objTestSetDictionary.Item("Not Completed Priority 3"))
    strNewText = Replace(strNewText, "NotCompletedPriority4", objTestSetDictionary.Item("Not Completed Priority 4"))
    strNewText = Replace(strNewText, "iP1NC", iP1NC)
    strNewText = Replace(strNewText, "iP2NC", iP2NC)
    strNewText = Replace(strNewText, "iP3NC", iP3NC)
    strNewText = Replace(strNewText, "iP4NC", iP4NC)
    
    '   Write the new text to the asp file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-" & "TestSetStatusStage1.txt", ForWriting, True)
    myFile.WriteLine strNewText
    myFile.Close

    '   Write out the data for the Progress percentage graph
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "TestPlannedProgressTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-" & "TestPlannedProgress.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "PassedPoints") > 0 Then
            '   Write the values
            myDest.WriteLine "<DCWC:DataPoint XValue=" & Chr(34) & "1" & Chr(34) & " YValues=" & Chr(34) & iPassedPercent & Chr(34) & " />"
        Else
            If InStr(1, rc, "FailedPoints") > 0 Then
                myDest.WriteLine "<DCWC:DataPoint XValue=" & Chr(34) & "1" & Chr(34) & " YValues=" & Chr(34) & iFailedPercent & Chr(34) & " />"
            Else
                If InStr(1, rc, "NotCompletedPoints") > 0 Then
                    myDest.WriteLine "<DCWC:DataPoint XValue=" & Chr(34) & "1" & Chr(34) & " YValues=" & Chr(34) & iNCPercent & Chr(34) & " />"
                Else
                    If InStr(1, rc, "NotRunPoints") > 0 Then
                         myDest.WriteLine "<DCWC:DataPoint XValue=" & Chr(34) & "1" & Chr(34) & " YValues=" & Chr(34) & iNoRunPercent & Chr(34) & " />"
                    Else
                        myDest.WriteLine rc
                    End If
                End If
            End If
        End If
    Loop
    
    '   Close the files
    myDest.Close
    mySource.Close
    
    '   Set blnDontRun
    blnDontRun = False

End Function
Public Function CountTestsRunByDate(ByVal strType As String, ByVal strDateFilter As String, Optional strStatus As String, Optional strPriority As String)
Dim tdcTestSetFactory As TestSetFactory
Dim tdcTestFilter As TDFilter
Dim intTestCount As Integer
Dim iAction As Integer
Dim rc As String
iAction = 0

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If
    
    '   See which kind of search we're doing
    If strSubProjectName <> "N/A" Then
        iAction = iAction + 1
    End If
    If strTestCycle <> "N/A" Then
        iAction = iAction + 2
    End If
    
    Set tdcTestSetFactory = tdc.TestSetFactory
    Set tdcTestFilter = tdcTestSetFactory.Filter
    tdcTestFilter.Filter(strProjectCycleLabel) = Chr(34) & strProjectName & Chr(34)
    tdcTestFilter.Filter(strTestPhaseCycleLabel) = Chr(34) & strTestPhase & Chr(34)
    Select Case iAction
        Case 1
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
        Case 2
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
        Case 3
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
    End Select
    
    Set colTestSets = tdcTestSetFactory.NewList(tdcTestFilter.Text)
    
    intTestCount = 0

    For Each objTestSet In colTestSets
        rc = objTestSet.TestSetFolder.Path
        If InStr(1, rc, "Unattached") = 0 And InStr(1, rc, "99. Archive") = 0 And InStr(1, rc, "Trash") = 0 Then
            intTestSetCount = ReturnTestSetCount(objTestSet, strType, strDateFilter, strStatus, strPriority)
            intTestCount = intTestCount + intTestSetCount
        End If
    Next
    
    CountTestsRunByDate = intTestCount
    
    '   Destroy objects
    Set tdcTestFilter = Nothing
    Set colTestSets = Nothing
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function ReturnTestSetCount(ByRef objTestSet, ByVal strType As String, ByVal strDateFilter As String, Optional strStatus As String, Optional strPriority As String)
Dim tdcTestFilter As TDFilter
Dim TSTestFact
    
    '   Set the TestSet Factory object
    Set TSTestFact = objTestSet.TSTestFactory
        '   Set up the filter
        Set tdcTestFilter = TSTestFact.Filter
            Select Case strType
                Case "Status"
                    Select Case strStatus
                        Case "No Run"
                                tdcTestFilter.Filter("TC_STATUS") = Chr(34) & strStatus & Chr(34)
                                tdcTestFilter.Filter("TC_PLAN_SCHEDULING_DATE") = strDateFilter
                        Case "No Run Daily", "No Run Weekly"
                                tdcTestFilter.Filter("TC_STATUS") = Chr(34) & "No Run" & Chr(34)
                                tdcTestFilter.Filter(strReplannedDateLabel) = strDateFilter
                        Case Else
                                tdcTestFilter.Filter("TC_STATUS") = Chr(34) & strStatus & Chr(34)
                                tdcTestFilter.Filter("TC_EXEC_DATE") = strDateFilter
                    End Select
                Case "Planned-Baseline"
                    tdcTestFilter.Filter("TC_PLAN_SCHEDULING_DATE") = strDateFilter
                    tdcTestFilter.Filter("TC_STATUS") = " Not (" & Chr(34) & "N/A" & Chr(34) & ")"
                Case "Replanned"
                    tdcTestFilter.Filter(strReplannedDateLabel) = strDateFilter
                    tdcTestFilter.Filter("TC_STATUS") = " Not (" & Chr(34) & "N/A" & Chr(34) & ")"
                Case "Actual"
                    tdcTestFilter.Filter("TC_STATUS") = " Not (" & Chr(34) & "N/A" & Chr(34) & " Or " & Chr(34) & "No Run" & Chr(34) & " Or " & Chr(34) & "Not Completed" & Chr(34) & ")"
                    tdcTestFilter.Filter("TC_EXEC_DATE") = strDateFilter
                Case "Status-Priority"
                    If strStatus = "No Run" Then
                        tdcTestFilter.Filter("TC_STATUS") = Chr(34) & strStatus & Chr(34)
                        tdcTestFilter.Filter("TC_PLAN_SCHEDULING_DATE") = strDateFilter
                    Else
                        tdcTestFilter.Filter("TC_STATUS") = Chr(34) & strStatus & Chr(34)
                        tdcTestFilter.Filter("TC_EXEC_DATE") = strDateFilter
                    End If
                    strMyPriority = GetFieldName("Plan: Priority", "TESTCYCL")
                    tdcTestFilter.Filter(strMyPriority) = Chr(34) & strPriority & Chr(34)
                Case "Planned-Priority"
                    tdcTestFilter.Filter("TC_PLAN_SCHEDULING_DATE") = strDateFilter
                    tdcTestFilter.Filter("TC_STATUS") = " Not (" & Chr(34) & "N/A" & Chr(34) & ")"
                    strMyPriority = GetFieldName("Plan: Priority", "TESTCYCL")
                    tdcTestFilter.Filter(strMyPriority) = Chr(34) & strPriority & Chr(34)
        End Select
                
        '   Get the list from the filter
        Set colTSTests = tdcTestFilter.NewList
                
        '   Return the count
        ReturnTestSetCount = colTSTests.Count
        
        '   Release objects
        Set TSTestFact = Nothing
        Set tdcTestFilter = Nothing
        Set colTSTests = Nothing
        
End Function
Public Sub GetTestsRun()
Dim intPlannedCount As Integer
Dim intActualCount As Integer
Dim intColumn As Integer
Dim strProjectNameCycleLabel As String
Dim strTestPhaseCycleLabel As String
Dim strDateFilter As String
Dim iCount As Integer
Dim arrWeekends
Dim myDate As Date
Dim dateBonkersDate As Date
Dim blnErrorPage As Boolean

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

iBase = 0
iRep = 0
iExe = 0
iPass = 0
iFail = 0
iNC = 0
iNR = 0
ReDim Preserve arrBaseline(iBase)
arrBaseline(iBase) = "Project Start|0"
ReDim Preserve arrReplanned(iRep)
arrReplanned(iRep) = "Project Start|0"
ReDim Preserve arrExecuted(iExe)
arrExecuted(iExe) = "Project Start|0"
ReDim Preserve arrPassed(iPass)
arrPassed(iPass) = "Project Start|0"
ReDim Preserve arrFailed(iFail)
arrFailed(iFail) = "Project Start|0"
ReDim Preserve arrNC(iNC)
arrNC(iNC) = "Project Start|0"
ReDim Preserve arrNR(iNR)
arrNR(iNR) = "Project Start|" & intTotalPlannedTests

    '   Get the bonkers date
    dateBonkersDate = DateAdd("yyyy", -2, TodaysDate)
   
    '   Get test set start and end dates
    dateBaselineStart = FindEarliestDate("Planned-Baseline")
    datePlannedStart = FindEarliestDate("Replanned")
    dateActualStart = FindEarliestDate("Actual")
    '   See if we've got no dates at all
    If dateBaselineStart = "00:00:00" And datePlannedStart = "00:00:00" And dateActualStart = "00:00:00" Then
        '   No tests sets set up or run so default to error page
        blnErrorPage = True
        GoTo WriteHTML
    End If
    '   If we've only got an actual date then use it
    If dateBaselineStart = "00:00:00" And datePlannedStart = "00:00:00" And dateActualStart <> "00:00:00" Then
        myStartDate = dateActualStart
    End If
    '   If we've only got a planned date then use it
    If dateBaselineStart = "00:00:00" And datePlannedStart <> "00:00:00" And dateActualStart = "00:00:00" Then
        '   See if we've got a bonkers date
        If datePlannedStart < dateBonkersDate Then
            myStartDate = TodaysDate
        Else
            myStartDate = datePlannedStart
        End If
    End If
    '   If we've only got a baseline date then use it
    If dateBaselineStart <> "00:00:00" And datePlannedStart = "00:00:00" And dateActualStart = "00:00:00" Then
        myStartDate = dateBaselineStart
    End If
    '   If we've got a baseline and a planned but no actual then work out which is earliest
    If dateBaselineStart <> "00:00:00" And datePlannedStart <> "00:00:00" And dateActualStart = "00:00:00" Then
        If datePlannedStart < dateBaselineStart Then
            '   See if we've got a bonkers date
            If datePlannedStart < dateBonkersDate Then
                myStartDate = dateBaselineStart
            Else
                myStartDate = datePlannedStart
            End If
        Else
            myStartDate = dateBaselineStart
        End If
    End If
    '   If we've got a baseline and an actual but no planned then work out which is earliest
    If dateBaselineStart <> "00:00:00" And datePlannedStart = "00:00:00" And dateActualStart <> "00:00:00" Then
        If dateActualStart < dateBaselineStart Then
            myStartDate = dateActualStart
        Else
            myStartDate = dateBaselineStart
        End If
    End If
    '   If we've got a planned and an actual but no baseline then work out which is earliest
    If dateBaselineStart = "00:00:00" And datePlannedStart <> "00:00:00" And dateActualStart <> "00:00:00" Then
        If dateActualStart < datePlannedStart Then
            myStartDate = dateActualStart
        Else
            '   See if we've got a bonkers date
            If datePlannedStart < dateBonkersDate Then
                myStartDate = dateActualStart
            Else
                myStartDate = datePlannedStart
            End If
        End If
    End If
    '   If we've got all of them then work out which is earliest
    If dateBaselineStart <> "00:00:00" And datePlannedStart <> "00:00:00" And dateActualStart <> "00:00:00" Then
        If dateActualStart < datePlannedStart And dateActualStart < dateBaselineStart Then
            myStartDate = dateActualStart
        Else
            If datePlannedStart < dateBaselineStart Then
                '   See if we've got a bonkers date
                If datePlannedStart < dateBonkersDate Then
                    myStartDate = dateBaselineStart
                Else
                    myStartDate = datePlannedStart
                End If
            Else
                myStartDate = dateBaselineStart
            End If
        End If
    End If
    
    '   Set the start date
    myDate = myStartDate
    
    '   Now get end dates
    dateBaselineEnd = FindLastDate("Planned-Baseline")
    datePlannedEnd = FindLastDate("Replanned")
    dateActualEnd = FindLastDate("Actual")
    '   See if we've got no baseline
    If dateBaselineEnd = "00:00:00" Then
        '   See if we've got no planned end
        If datePlannedEnd = "00:00:00" Then
            '   See if we've got no actual
            If dateActualEnd = "00:00:00" Then
                '   Default to today
                myEndDate = TodaysDate
            Else
                If dateActualEnd < TodaysDate Then
                    myEndDate = TodaysDate
                Else
                    myEndDate = dateActualEnd
                End If
            End If
        Else
            '   See if we've got no actual
            If dateActualEnd = "00:00:00" Then
                If datePlannedEnd < TodaysDate Then
                    myEndDate = TodaysDate
                Else
                    myEndDate = datePlannedEnd
                End If
            Else
                '   See if actual is after planned
                If dateActualEnd > datePlannedEnd Then
                    If dateActualEnd < TodaysDate Then
                        myEndDate = TodaysDate
                    Else
                        myEndDate = dateActualEnd
                    End If
                Else
                    If datePlannedEnd < TodaysDate Then
                        myEndDate = TodaysDate
                    Else
                        myEndDate = datePlannedEnd
                    End If
                End If
            End If
        End If
    Else
        '   See if we've got no planned end
        If datePlannedEnd = "00:00:00" Then
            '   See if we've got no actual
            If dateActualEnd = "00:00:00" Then
                '   Default to today if baseline end is less
                If dateBaselineEnd < TodaysDate Then
                    myEndDate = TodaysDate
                Else
                    myEndDate = dateBaselineEnd
                End If
            Else
                '   See if actual is after baseline
                If dateActualEnd > dateBaselineEnd Then
                    If dateActualEnd < TodaysDate Then
                        myEndDate = TodaysDate
                    Else
                        myEndDate = dateActualEnd
                    End If
                Else
                    If dateBaselineEnd < TodaysDate Then
                        myEndDate = TodaysDate
                    Else
                        myEndDate = dateBaselineEnd
                    End If
                End If
            End If
        Else
            '   See if we've got no actual
            If dateActualEnd = "00:00:00" Then
                '   See if planned is after baseline
                If datePlannedEnd > dateBaselineEnd Then
                    If datePlannedEnd < TodaysDate Then
                        myEndDate = TodaysDate
                    Else
                        myEndDate = datePlannedEnd
                    End If
                Else
                    If dateBaselineEnd < Today Then
                        myEndDate = TodaysDate
                    Else
                        myEndDate = dateBaselineEnd
                    End If
                End If
            Else
                '   See if actual is after planned
                If dateActualEnd > datePlannedEnd Then
                    '   See if the actual is after the baseline
                    If dateActualEnd > dateBaselineEnd Then
                        If dateActualEnd < TodaysDate Then
                            myEndDate = TodaysDate
                        Else
                            myEndDate = dateActualEnd
                        End If
                    Else
                        If dateBaselineEnd < TodaysDate Then
                            myEndDate = TodaysDate
                        Else
                            myEndDate = dateBaselineEnd
                        End If
                    End If
                Else
                    '   See if the planned is after the baseline
                    If datePlannedEnd > dateBaselineEnd Then
                        If datePlannedEnd < TodaysDate Then
                            myEndDate = TodaysDate
                        Else
                            myEndDate = datePlannedEnd
                        End If
                    Else
                        If dateBaselineEnd < Today Then
                            myEndDate = TodaysDate
                        Else
                            myEndDate = dateBaselineEnd
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    '   Set the end date
    dateEndDate = myEndDate
       
    '   Find the last weeks set of days range
    dateLastWeek = DateAdd("d", -7, TodaysDate)
    '   Find the previous months start date
    datePreviousMonth = DateAdd("m", -1, dateLastWeek)
    If Weekday(datePreviousMonth) = 1 Then
        datePreviousMonth = DateAdd("d", 1, datePreviousMonth)
    Else
        If Weekday(datePreviousMonth) = 7 Then
            datePreviousMonth = DateAdd("d", 2, datePreviousMonth)
        End If
    End If
    '   Find the previous months
    arrMonths = FindMonths(myDate, datePreviousMonth)
    If arrMonths(0) <> "No Month" Then
        arrWeekends = FindWeekends(arrMonths(UBound(arrMonths)), dateLastWeek)
    Else
        arrWeekends = FindWeekends(myDate, dateLastWeek)
    End If
    
    '   Write out the month values
    If arrMonths(0) <> "No Month" Then
        myDate = WriteOutMonthsWeeks(arrMonths, myDate)
    End If
    '   Write out the week values
    If arrWeekends(0) <> "No Weeks" Then
        myDate = WriteOutMonthsWeeks(arrWeekends, myDate)
    End If
    
    '   Do the daily values upto todays date
    Do While myDate <= TodaysDate
        
        strDateFilter = myDate
        
        intBaselineCount = 0
        intBaselineCount = CountTestsRunByDate("Planned-Baseline", strDateFilter)
        
        intPlannedCount = 0
        intPlannedCount = CountTestsRunByDate("Replanned", strDateFilter)
    
        intActualCount = 0
        intActualCount = CountTestsRunByDate("Actual", strDateFilter)
        
        intPassedCount = 0
        intPassedCount = CountTestsRunByDate("Status", strDateFilter, "Passed")

        intFailedCount = 0
        intFailedCount = CountTestsRunByDate("Status", strDateFilter, "Failed")

        intNCCount = 0
        intNCCount = CountTestsRunByDate("Status", strDateFilter, "Not Completed")
        
        intNRCount = 0
         
        '   Ignore if we're on a weekend with no values
        If (Weekday(strDateFilter) = 7 Or Weekday(strDateFilter) = 1) _
            And intPlannedCount = 0 _
            And intActualCount = 0 _
            And intPassedCount = 0 _
            And intFailedCount = 0 _
            And intNCCount = 0 _
            Then
            GoTo ExitLoop
        End If
        
        '   Add to the arrays
        iBase = iBase + 1
        ReDim Preserve arrBaseline(iBase)
        arrBaseline(iBase) = strDateFilter & "|" & intBaselineCount
        iRep = iRep + 1
        ReDim Preserve arrReplanned(iRep)
        arrReplanned(iRep) = strDateFilter & "|" & intPlannedCount
        iExe = iExe + 1
        ReDim Preserve arrExecuted(iExe)
        arrExecuted(iExe) = strDateFilter & "|" & intActualCount
        iPass = iPass + 1
        ReDim Preserve arrPassed(iPass)
        arrPassed(iPass) = strDateFilter & "|" & intPassedCount
        iFail = iFail + 1
        ReDim Preserve arrFailed(iFail)
        arrFailed(iFail) = strDateFilter & "|" & intFailedCount
        iNC = iNC + 1
        ReDim Preserve arrNC(iNC)
        arrNC(iNC) = strDateFilter & "|" & intNCCount
        iNR = iNR + 1
        ReDim Preserve arrNR(iNR)
        
        '   Make these cumulative
        If iBase > 0 Then
            myoldsplit = Split(arrBaseline(iBase - 1), "|")
            mycurrsplit = Split(arrBaseline(iBase), "|")
            iBaseTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrBaseline(iBase) = mycurrsplit(0) & "|" & iBaseTot
        End If
        If iRep > 0 Then
            myoldsplit = Split(arrReplanned(iRep - 1), "|")
            mycurrsplit = Split(arrReplanned(iRep), "|")
            iRepTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrReplanned(iRep) = mycurrsplit(0) & "|" & iRepTot
        End If
        If iExe > 0 Then
            myoldsplit = Split(arrExecuted(iExe - 1), "|")
            mycurrsplit = Split(arrExecuted(iExe), "|")
            iExeTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrExecuted(iExe) = mycurrsplit(0) & "|" & iExeTot
        End If
        If iPass > 0 Then
            myoldsplit = Split(arrPassed(iPass - 1), "|")
            mycurrsplit = Split(arrPassed(iPass), "|")
            iPassTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrPassed(iPass) = mycurrsplit(0) & "|" & iPassTot
        End If
        If iFail > 0 Then
            myoldsplit = Split(arrFailed(iFail - 1), "|")
            mycurrsplit = Split(arrFailed(iFail), "|")
            iFailTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrFailed(iFail) = mycurrsplit(0) & "|" & iFailTot
        End If
        If iNC > 0 Then
            myoldsplit = Split(arrNC(iNC - 1), "|")
            mycurrsplit = Split(arrNC(iNC), "|")
            iNCTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrNC(iNC) = mycurrsplit(0) & "|" & iNCTot
        End If
        If iNR > 0 Then
            iNRTot = intTotalPlannedTests - iExeTot
            arrNR(iNR) = mycurrsplit(0) & "|" & iNRTot
        End If
ExitLoop:
        myDate = myDate + 1
    Loop
    
    '   Now get just the planned values out into the future if required
    If dateEndDate > TodaysDate Then
        iDays = DateDiff("d", TodaysDate, dateEndDate)
        If iDays <= 30 Then
            myDate = WriteDailyValues(DateAdd("d", 1, TodaysDate), dateEndDate, "Y")
            GoTo WriteHTML
        End If
        '   Get the date next week
        dateNextWeek = DateAdd("d", 7, TodaysDate)
        '   See if it's past our end date
        If dateNextWeek >= dateEndDate Then
            myDate = WriteDailyValues(DateAdd("d", 1, TodaysDate), dateEndDate, "Y")
            GoTo WriteHTML
        End If
        dateNextMonth = DateAdd("m", 1, dateNextWeek)
        '   See if it's past our end date
        If dateNextMonth >= dateEndDate Then
            arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
            If arrWeekends(0) <> "No Weeks" Then
                If arrWeekends(0) < dateEndDate Then
                    myDate = WriteDailyValues(DateAdd("d", 1, TodaysDate), dateNextWeek, "Y")
                    myDate = WriteOutMonthsWeeks(arrWeekends, myDate, "Y")
                Else
                    myDate = WriteDailyValues(DateAdd("d", 1, TodaysDate), dateEndDate, "Y")
                End If
                If myDate < dateEndDate Then
                    myDate = WriteDailyValues(myDate, dateEndDate, "Y")
                End If
            Else
                myDate = WriteDailyValues(myDate, dateEndDate, "Y")
            End If
        End If
        '   See how many days between dateNextMonth and end date
        iDays = DateDiff("d", dateNextMonth, dateEndDate)
        '   If it's less than 30 get weeks else get months
        If iDays > 0 And iDays <= 30 Then
            arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
            myDate = WriteDailyValues(myDate, dateNextWeek, "Y")
            myDate = WriteOutMonthsWeeks(arrWeekends, myDate, "Y", "Y")
        Else
            If iDays < 0 Then
                GoTo WriteHTML
            End If
            myDate = WriteDailyValues(myDate, dateNextWeek, "Y")
            arrWeekends = FindWeekends(myDate, dateNextMonth)
            myDate = WriteOutMonthsWeeks(arrWeekends, myDate, "Y", "N")
            arrMonths = FindMonths(myDate, dateEndDate)
            myDate = WriteOutMonthsWeeks(arrMonths, myDate, "Y", "Y")
        End If
        
    End If
WriteHTML:
    '   See if we've got to produce the error page or not
    If blnErrorPage = False Then
    
        '   Now write out the PlannedvsExecuted table info
        fso.CopyFile strTemplatePath & "TestPlannedvsExecutedTableTemplate.txt", strFolderPath & strPathandFileName & "-TestPlannedvsExecutedTable.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestPlannedvsExecutedTable.asp", ForAppending, True)
        
        '   Write out the table header details etc
        myFile.WriteLine "<table border=1 align=center>"
        myFile.WriteLine "<tr><th colspan=7 align=center>" & strHeader & "</th></tr>"
        myFile.WriteLine "<tr bgcolor='#56A5EC'><td>&nbsp;</td><td>Baseline Planned</td><td>Replanned</td><td>Executed</td><td>Passed</td><td>Failed</td><td>Not Completed</td></tr>"
        iBaselineCount = UBound(arrBaseline)
        iPassedCount = UBound(arrPassed)
        If iBaselineCount = iPassedCount Then
            iTotal = iBaselineCount
        Else
            iTotal = iPassedCount
        End If
        i = 0
        Do
            '   Get the values from each of the arrays for this array element
            aSplit = Split(arrBaseline(i), "|")
            bSplit = Split(arrReplanned(i), "|")
            cSplit = Split(arrPassed(i), "|")
            dSplit = Split(arrFailed(i), "|")
            eSplit = Split(arrNC(i), "|")
            
            '   If we're on the first element then just default to project start and zeros
            If i = 0 Then
                myFile.WriteLine "<tr bgcolor ='#B4CFEC'><td>Project Start</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td></tr>"
            Else
                If IsEven(i) = True Then
                    myFile.WriteLine ("<tr bgcolor ='#B4CFEC'>")
                Else
                    myFile.WriteLine ("<tr bgcolor ='ivory'>")
                End If
                '   Write a row to the table
                myFile.WriteLine "<td>" & aSplit(0) & "</td><td>" & aSplit(1) & "</td><td>" & bSplit(1) & "</td><td>" & CInt(cSplit(1)) + CInt(dSplit(1)) & "</td><td>" & cSplit(1) & "</td><td>" & dSplit(1) & "</td><td>" & eSplit(1) & "</td></tr>"
            End If
            i = i + 1
        Loop Until i > iTotal
        If iBaselineCount > iPassedCount Then
            For j = i To UBound(arrBaseline)
                If IsEven(i) = True Then
                    myFile.WriteLine ("<tr bgcolor ='#B4CFEC'>")
                Else
                    myFile.WriteLine ("<tr bgcolor ='ivory'>")
                End If
                aSplit = Split(arrBaseline(j), "|")
                bSplit = Split(arrReplanned(j), "|")
                '   Write a row to the table
                myFile.WriteLine "<td>" & aSplit(0) & "</td><td>" & aSplit(1) & "</td><td>" & bSplit(1) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
            Next
        End If
        '   Write the remainder of the html
        myFile.WriteLine "</table></body></html>"
        '   Close the file
        myFile.Close

        '   Now write out the PlannedvsExecuted info
        Set mySource = fso.OpenTextFile(strTemplatePath & "TestPlannedvsExecutedTemplate.txt", ForReading)
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestPlannedvsExecuted.aspx", ForWriting, True)
        Do While mySource.AtEndOfStream <> True
            rc = mySource.ReadLine
            If InStr(1, rc, "BaselinePlannedPoints") > 0 Then
                '   Loop round our Baseline array
                For i = 0 To UBound(arrBaseline)
                    '   Split the file
                    mySplit = Split(arrBaseline(i), "|")
                    '   Re-format the date part
                    If mySplit(0) <> "Project Start" Then
                        TheDate = Format(mySplit(0), "dd mmm yy")
                    Else
                        TheDate = mySplit(0)
                    End If
                    '   Write the value
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & TheDate & Chr(34) & " />"
                Next
            Else
                If InStr(1, rc, "ReplannedPoints") > 0 Then
                    '   Loop round our Replanned array
                    For i = 0 To UBound(arrReplanned)
                        '   Split the file
                        mySplit = Split(arrReplanned(i), "|")
                        '   Re-format the date part
                        If mySplit(0) <> "Project Start" Then
                            TheDate = Format(mySplit(0), "dd mmm yy")
                        Else
                            TheDate = mySplit(0)
                        End If
                        '   Write the value
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & TheDate & Chr(34) & " />"
                    Next
                Else
                    If InStr(1, rc, "ExecutedPoints") > 0 Then
                        '   Loop round our Executed array
                        For i = 0 To UBound(arrExecuted)
                            '   Split the file
                            mySplit = Split(arrExecuted(i), "|")
                            '   Re-format the date part
                            If mySplit(0) <> "Project Start" Then
                                TheDate = Format(mySplit(0), "dd mmm yy")
                            Else
                                TheDate = mySplit(0)
                            End If
                            '   Write the value
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & TheDate & Chr(34) & " />"
                        Next
                    Else
                        If InStr(1, rc, "PassedPoints") > 0 Then
                            '   Loop round our passed array
                            For i = 0 To UBound(arrPassed)
                                '   Split the file
                                mySplit = Split(arrPassed(i), "|")
                                '   Re-format the date part
                                If mySplit(0) <> "Project Start" Then
                                    TheDate = Format(mySplit(0), "dd mmm yy")
                                Else
                                    TheDate = mySplit(0)
                                End If
                                '   Write the value
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & TheDate & Chr(34) & " />"
                            Next
                        Else
                            If InStr(1, rc, "FailedPoints") > 0 Then
                                '   Loop round our failed array
                                For i = 0 To UBound(arrFailed)
                                    '   Split the file
                                    mySplit = Split(arrFailed(i), "|")
                                    '   Re-format the date part
                                    If mySplit(0) <> "Project Start" Then
                                        TheDate = Format(mySplit(0), "dd mmm yy")
                                    Else
                                        TheDate = mySplit(0)
                                    End If
                                    '   Write the value
                                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & TheDate & Chr(34) & " />"
                                Next
                            Else
                                If InStr(1, rc, "NotCompletedPoints") > 0 Then
                                    '   Loop round our Not completed array
                                    For i = 0 To UBound(arrNC)
                                        '   Split the file
                                        mySplit = Split(arrNC(i), "|")
                                        '   Re-format the date part
                                        If mySplit(0) <> "Project Start" Then
                                            TheDate = Format(mySplit(0), "dd mmm yy")
                                        Else
                                            TheDate = mySplit(0)
                                        End If
                                        '   Write the value
                                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & TheDate & Chr(34) & " />"
                                    Next
                                Else
                                    myDest.WriteLine rc
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Loop
        
        '   Close the files
        myDest.Close
        mySource.Close
        
        '   Update the main test set file with the correct path for this image
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", ForReading)
        strText = myDest.ReadAll
        myDest.Close
        
        '   Change the value
        strText = Replace(strText, "TestPlannedvsExecuted", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestPlannedvsExecuted.aspx" & Chr(34))
        
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", ForWriting, True)
        myDest.WriteLine strText
        myDest.Close
        
    Else
        '   See if test set status file exists
        If fso.FileExists(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt") = True Then
            '   Delete it
            fso.DeleteFile strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", True
        End If
        '   Create it again, from the error template
        fso.CopyFile strTemplatePath & "TestSetsMissingTemplate.txt", strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt"
    
        '   Open the file to read the data
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        '   Change the header
        strText = Replace(strText, "strHeader", "No Test Script data for " & strHeader)
    
        '   Write the data outr
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", ForWriting, True)
        myFile.WriteLine strText
        myFile.Close
    
        '   Set the don't run flag
        blnDontRun = True
            
    End If
    
Exit Sub

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Sub
Public Function WriteOutMonthsWeeks(ByVal arrArray As Variant, ByVal dateThisStartDate As Date, Optional strPlannedOnly As String, Optional strEnd As String, Optional strDateOnly As String)
 
    If arrArray(0) <> "No Month" And arrArray(0) <> "No Weeks" Then
    
        If strDateOnly = "Y" Then
            
            For Each Ele In arrArray
                dateThisStartDate = Ele
            Next
            WriteOutMonthsWeeks = dateThisStartDate
            Exit Function
        End If
    
        For Each Ele In arrArray
            
            intBaselineCount = 0
            intPlannedCount = 0
            intActualCount = 0
            intPassedCount = 0
            intFailedCount = 0
            intNCCount = 0
            intNRCount = 0
            
            strDateFilter = ">= " & dateThisStartDate & " And < " & Ele
            
            '   See if we're doing planned only
            If strPlannedOnly <> "Y" Then
                intBaselineCount = CountTestsRunByDate("Planned-Baseline", strDateFilter)
                intPlannedCount = CountTestsRunByDate("Replanned", strDateFilter)
                intActualCount = CountTestsRunByDate("Actual", strDateFilter)
                intPassedCount = CountTestsRunByDate("Status", strDateFilter, "Passed")
                intFailedCount = CountTestsRunByDate("Status", strDateFilter, "Failed")
                intNCCount = CountTestsRunByDate("Status", strDateFilter, "Not Completed")
                intNRCount = 0
            
                '   Add to the arrays
                iBase = iBase + 1
                ReDim Preserve arrBaseline(iBase)
                arrBaseline(iBase) = dateThisStartDate & "|" & intBaselineCount
                iRep = iRep + 1
                ReDim Preserve arrReplanned(iRep)
                arrReplanned(iRep) = dateThisStartDate & "|" & intPlannedCount
                iExe = iExe + 1
                ReDim Preserve arrExecuted(iExe)
                arrExecuted(iExe) = dateThisStartDate & "|" & intActualCount
                iPass = iPass + 1
                ReDim Preserve arrPassed(iPass)
                arrPassed(iPass) = dateThisStartDate & "|" & intPassedCount
                iFail = iFail + 1
                ReDim Preserve arrFailed(iFail)
                arrFailed(iFail) = dateThisStartDate & "|" & intFailedCount
                iNC = iNC + 1
                ReDim Preserve arrNC(iNC)
                arrNC(iNC) = dateThisStartDate & "|" & intNCCount
                iNR = iNR + 1
                ReDim Preserve arrNR(iNR)
                
                '   Make these cumulative
                If iBase > 0 Then
                    myoldsplit = Split(arrBaseline(iBase - 1), "|")
                    mycurrsplit = Split(arrBaseline(iBase), "|")
                    iBaseTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrBaseline(iBase) = mycurrsplit(0) & "|" & iBaseTot
                End If
                If iRep > 0 Then
                    myoldsplit = Split(arrReplanned(iRep - 1), "|")
                    mycurrsplit = Split(arrReplanned(iRep), "|")
                    iRepTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrReplanned(iRep) = mycurrsplit(0) & "|" & iRepTot
                End If
                If iExe > 0 Then
                    myoldsplit = Split(arrExecuted(iExe - 1), "|")
                    mycurrsplit = Split(arrExecuted(iExe), "|")
                    iExeTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrExecuted(iExe) = mycurrsplit(0) & "|" & iExeTot
                End If
                If iPass > 0 Then
                    myoldsplit = Split(arrPassed(iPass - 1), "|")
                    mycurrsplit = Split(arrPassed(iPass), "|")
                    iPassTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrPassed(iPass) = mycurrsplit(0) & "|" & iPassTot
                End If
                If iFail > 0 Then
                    myoldsplit = Split(arrFailed(iFail - 1), "|")
                    mycurrsplit = Split(arrFailed(iFail), "|")
                    iFailTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrFailed(iFail) = mycurrsplit(0) & "|" & iFailTot
                End If
                If iNC > 0 Then
                    myoldsplit = Split(arrNC(iNC - 1), "|")
                    mycurrsplit = Split(arrNC(iNC), "|")
                    iNCTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrNC(iNC) = mycurrsplit(0) & "|" & iNCTot
                End If
                If iNR > 0 Then
                    iNRTot = intTotalPlannedTests - iExeTot
                    arrNR(iNR) = mycurrsplit(0) & "|" & iNRTot
                End If
            Else
                intBaselineCount = CountTestsRunByDate("Planned-Baseline", strDateFilter)
                intPlannedCount = CountTestsRunByDate("Replanned", strDateFilter)
            
                iBase = iBase + 1
                ReDim Preserve arrBaseline(iBase)
                arrBaseline(iBase) = dateThisStartDate & "|" & intBaselineCount
                iRep = iRep + 1
                ReDim Preserve arrReplanned(iRep)
                arrReplanned(iRep) = dateThisStartDate & "|" & intPlannedCount
                iExe = iExe + 1
            
                '   Make these cumulative
                If iBase > 0 Then
                    myoldsplit = Split(arrBaseline(iBase - 1), "|")
                    mycurrsplit = Split(arrBaseline(iBase), "|")
                    iBaseTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrBaseline(iBase) = mycurrsplit(0) & "|" & iBaseTot
                End If
                If iRep > 0 Then
                    myoldsplit = Split(arrReplanned(iRep - 1), "|")
                    mycurrsplit = Split(arrReplanned(iRep), "|")
                    iRepTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrReplanned(iRep) = mycurrsplit(0) & "|" & iRepTot
                End If
            End If
            dateThisStartDate = Ele
            
        Next
    End If
    
    '   If doing planned only check the following
    If strPlannedOnly = "Y" Then
        '   See if we're ending
        If strEnd = "Y" Then
            '   See if we're under the end date
            If dateThisStartDate <= dateEndDate Then
                intPlannedCount = 0
                
                strDateFilter = ">= " & dateThisStartDate & " And <= " & dateEndDate
                
                intBaselineCount = CountTestsRunByDate("Planned-Baseline", strDateFilter)
                intPlannedCount = CountTestsRunByDate("Replanned", strDateFilter)
                
                iBase = iBase + 1
                ReDim Preserve arrBaseline(iBase)
                arrBaseline(iBase) = dateThisStartDate & "|" & intBaselineCount
                iRep = iRep + 1
                ReDim Preserve arrReplanned(iRep)
                arrReplanned(iRep) = dateThisStartDate & "|" & intPlannedCount
            
                '   Make these cumulative
                If iBase > 0 Then
                    myoldsplit = Split(arrBaseline(iBase - 1), "|")
                    mycurrsplit = Split(arrBaseline(iBase), "|")
                    iBaseTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrBaseline(iBase) = mycurrsplit(0) & "|" & iBaseTot
                End If
                If iRep > 0 Then
                    myoldsplit = Split(arrReplanned(iRep - 1), "|")
                    mycurrsplit = Split(arrReplanned(iRep), "|")
                    iRepTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                    arrReplanned(iRep) = mycurrsplit(0) & "|" & iRepTot
                End If
            End If
            
            dateThisStartDate = dateEndDate
        End If
    End If
    
    '   Return the date for the next section
    WriteOutMonthsWeeks = dateThisStartDate
    
End Function
Public Function WriteDailyValues(ByVal myDate As Date, ByVal myEndDate As Date, Optional strPlannedOnly As String)

    
    '   Do the daily values upto todays date
    Do While myDate <= myEndDate
        
        strDateFilter = myDate
        
        intBaselineCount = 0
        intBaselineCount = CountTestsRunByDate("Planned-Baseline", strDateFilter)
        
        intPlannedCount = 0
        intPlannedCount = CountTestsRunByDate("Replanned", strDateFilter)
        
        '   See if we're doing planned only
        If strPlannedOnly <> "Y" Then
            intActualCount = 0
            intActualCount = CountTestsRunByDate("Actual", strDateFilter)
            
            intPassedCount = 0
            intPassedCount = CountTestsRunByDate("Status", strDateFilter, "Passed")
    
            intFailedCount = 0
            intFailedCount = CountTestsRunByDate("Status", strDateFilter, "Failed")
    
            intNCCount = 0
            intNCCount = CountTestsRunByDate("Status", strDateFilter, "Not Completed")
            
            intNRCount = 0
            'intNRCount = CountTestsRunByDate("Status", strDateFilter, "No Run")
                
            If (Weekday(strDateFilter) = 7 Or Weekday(strDateFilter) = 1) _
                And intPlannedCount = 0 _
                And intActualCount = 0 _
                And intPassedCount = 0 _
                And intFailedCount = 0 _
                And intNCCount = 0 _
                Then
                GoTo ExitLoop
            End If
            
            '   Add to the arrays
            iBase = iBase + 1
            ReDim Preserve arrBaseline(iBase)
            arrBaseline(iBase) = strDateFilter & "|" & intBaselineCount
            iRep = iRep + 1
            ReDim Preserve arrReplanned(iRep)
            arrReplanned(iRep) = strDateFilter & "|" & intPlannedCount
            iExe = iExe + 1
            ReDim Preserve arrExecuted(iExe)
            arrExecuted(iExe) = strDateFilter & "|" & intActualCount
            iPass = iPass + 1
            ReDim Preserve arrPassed(iPass)
            arrPassed(iPass) = strDateFilter & "|" & intPassedCount
            iFail = iFail + 1
            ReDim Preserve arrFailed(iFail)
            arrFailed(iFail) = strDateFilter & "|" & intFailedCount
            iNC = iNC + 1
            ReDim Preserve arrNC(iNC)
            arrNC(iNC) = strDateFilter & "|" & intNCCount
            iNR = iNR + 1
            ReDim Preserve arrNR(iNR)
            
            '   Make these cumulative
            If iBase > 0 Then
                myoldsplit = Split(arrBaseline(iBase - 1), "|")
                mycurrsplit = Split(arrBaseline(iBase), "|")
                iBaseTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrBaseline(iBase) = mycurrsplit(0) & "|" & iBaseTot
            End If
            If iRep > 0 Then
                myoldsplit = Split(arrReplanned(iRep - 1), "|")
                mycurrsplit = Split(arrReplanned(iRep), "|")
                iRepTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrReplanned(iRep) = mycurrsplit(0) & "|" & iRepTot
            End If
            If iExe > 0 Then
                myoldsplit = Split(arrExecuted(iExe - 1), "|")
                mycurrsplit = Split(arrExecuted(iExe), "|")
                iExeTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrExecuted(iExe) = mycurrsplit(0) & "|" & iExeTot
            End If
            If iPass > 0 Then
                myoldsplit = Split(arrPassed(iPass - 1), "|")
                mycurrsplit = Split(arrPassed(iPass), "|")
                iPassTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrPassed(iPass) = mycurrsplit(0) & "|" & iPassTot
            End If
            If iFail > 0 Then
                myoldsplit = Split(arrFailed(iFail - 1), "|")
                mycurrsplit = Split(arrFailed(iFail), "|")
                iFailTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrFailed(iFail) = mycurrsplit(0) & "|" & iFailTot
            End If
            If iNC > 0 Then
                myoldsplit = Split(arrNC(iNC - 1), "|")
                mycurrsplit = Split(arrNC(iNC), "|")
                iNCTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrNC(iNC) = mycurrsplit(0) & "|" & iNCTot
            End If
            If iNR > 0 Then
                iNRTot = intTotalPlannedTests - iExeTot
                arrNR(iNR) = mycurrsplit(0) & "|" & iNRTot
            End If
        Else
            If (Weekday(strDateFilter) = 7 Or Weekday(strDateFilter) = 1) _
                And intPlannedCount = 0 And intBaselineCount = 0 _
                Then
                GoTo ExitLoop
            End If
            
            '   Add to the arrays
            iBase = iBase + 1
            ReDim Preserve arrBaseline(iBase)
            arrBaseline(iBase) = strDateFilter & "|" & intBaselineCount
            iRep = iRep + 1
            ReDim Preserve arrReplanned(iRep)
            arrReplanned(iRep) = strDateFilter & "|" & intPlannedCount
        
            '   Make these cumulative
            If iBase > 0 Then
                myoldsplit = Split(arrBaseline(iBase - 1), "|")
                mycurrsplit = Split(arrBaseline(iBase), "|")
                iBaseTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrBaseline(iBase) = mycurrsplit(0) & "|" & iBaseTot
            End If
            If iRep > 0 Then
                myoldsplit = Split(arrReplanned(iRep - 1), "|")
                mycurrsplit = Split(arrReplanned(iRep), "|")
                iRepTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrReplanned(iRep) = mycurrsplit(0) & "|" & iRepTot
            End If
        End If
            
ExitLoop:
        myDate = myDate + 1
    Loop
    
    WriteDailyValues = myDate
    
End Function
Public Function FindEarliestDate(ByVal strType As String)
Dim tdcTestSetFactory As TestSetFactory
Dim tdcTestFilter As TDFilter
Dim iAction As Integer
Dim TheDate As Date
Dim myTemp As Date
iAction = 0

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    '   See which kind of search we're doing
    If strSubProjectName <> "N/A" Then
        iAction = iAction + 1
    End If
    If strTestCycle <> "N/A" Then
        iAction = iAction + 2
    End If
    
    '   Set up the test set filter
    Set tdcTestSetFactory = tdc.TestSetFactory
    Set tdcTestFilter = tdcTestSetFactory.Filter
    tdcTestFilter.Filter(strProjectCycleLabel) = Chr(34) & strProjectName & Chr(34)
    tdcTestFilter.Filter(strTestPhaseCycleLabel) = Chr(34) & strTestPhase & Chr(34)
    Select Case iAction
        Case 1
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
        Case 2
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
        Case 3
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
    End Select
    
    Set colTestSets = tdcTestSetFactory.NewList(tdcTestFilter.Text)
    
    intTestCount = 0
    
    For Each objTestSet In colTestSets
        rc = objTestSet.TestSetFolder.Path
        If InStr(1, rc, "Unattached") = 0 And InStr(1, rc, "99. Archive") = 0 And InStr(1, rc, "Trash") = 0 Then
                TheDate = GetEarliestLastDate(objTestSet, strType, "Early")
                If myTemp = "00:00:00" Then
                        myTemp = TheDate
                Else
                        If TheDate <> "00:00:00" Then
                                If TheDate < myTemp Then
                                        myTemp = TheDate
                                End If
                        End If
                End If
        End If
    Next
    
    '   Return the earliest date for the phase
    FindEarliestDate = myTemp
    
    '   Release objects
    Set tdcTestFilter = Nothing
    Set colTestSets = Nothing

    Exit Function
    
ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
    
End Function
Public Function GetEarliestLastDate(ByRef objTestSet, ByVal strType As String, ByVal strEarlyLast As String)
Dim objTSTestFactory
Dim tdcTestFilter
Dim myDate As Date
Dim myTemp As Date

    '   Set up test factory
    Set objTSTestFactory = objTestSet.TSTestFactory
    '   Create a filter
    Set tdcTestFilter = objTSTestFactory.Filter
    '   Get the rest by type
    Select Case strType
        Case "Planned-Baseline"
            tdcTestFilter.Filter("TC_PLAN_SCHEDULING_DATE") = "Not " & Chr(34) & Chr(34)
        Case "Replanned"
            tdcTestFilter.Filter(strReplannedDateLabel) = "Not " & Chr(34) & Chr(34)
        Case Else
            tdcTestFilter.Filter("TC_EXEC_DATE") = "Not " & Chr(34) & Chr(34)
    End Select
    Set colTSTests = tdcTestFilter.NewList
    For Each Ele In colTSTests
        Select Case strType
            Case "Planned-Baseline"
                myDate = Ele.Field("TC_PLAN_SCHEDULING_DATE")
            Case "Replanned"
                myDate = Ele.Field(strReplannedDateLabel)
            Case Else
                myDate = Ele.Field("TC_EXEC_DATE")
        End Select
        If myTemp = "00:00:00" Then
            myTemp = myDate
        Else
            If strEarlyLast = "Early" Then
                If myDate < myTemp Then
                    myTemp = myDate
                End If
            Else
                If myDate > myTemp Then
                    myTemp = myDate
                End If
            End If
        End If
    Next
    
    GetEarliestLastDate = myTemp
    
    '   Release objects
    Set objTSTestFactory = Nothing
    Set tdcTestFilter = Nothing
    Set colTSTests = Nothing

End Function
Public Function FindLastDate(ByVal strType As String)
Dim tdcTestSetFactory
Dim tdcTestFilter
Dim iAction As Integer
Dim TheDate As Date
Dim myTemp As Date
iAction = 0

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    '   See which kind of search we're doing
    If strSubProjectName <> "N/A" Then
        iAction = iAction + 1
    End If
    If strTestCycle <> "N/A" Then
        iAction = iAction + 2
    End If
    
    '   Set up the test set filter
    Set tdcTestSetFactory = tdc.TestSetFactory
    Set tdcTestFilter = tdcTestSetFactory.Filter
    tdcTestFilter.Filter(strProjectCycleLabel) = Chr(34) & strProjectName & Chr(34)
    tdcTestFilter.Filter(strTestPhaseCycleLabel) = Chr(34) & strTestPhase & Chr(34)
    Select Case iAction
        Case 1
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
        Case 2
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
        Case 3
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
    End Select
    
    Set colTestSets = tdcTestSetFactory.NewList(tdcTestFilter.Text)
    
    intTestCount = 0
    
    For Each objTestSet In colTestSets
        rc = objTestSet.TestSetFolder.Path
        If InStr(1, rc, "Unattached") = 0 And InStr(1, rc, "99. Archive") = 0 And InStr(1, rc, "Trash") = 0 Then
                TheDate = GetEarliestLastDate(objTestSet, strType, "Late")
                If myTemp = "00:00:00" Then
                        myTemp = TheDate
                Else
                        If TheDate <> "00:00:00" Then
                                If TheDate > myTemp Then
                                        myTemp = TheDate
                                End If
                        End If
                End If
        End If
    Next
    
    '   Return the earliest date for the phase
    FindLastDate = myTemp
    
    '   Release objects
    Set tdcTestFilter = Nothing
    Set colTestSets = Nothing

    Exit Function
    
ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
    
End Function
Public Function CountTestsByFunctionalArea()
Dim tdcTestSetFactory
Dim tdcTestFilter
Dim intTestCount As Integer
Dim iAction As Integer
Dim rc As String
iAction = 0

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    '   See which kind of search we're doing
    If strSubProjectName <> "N/A" Then
        iAction = iAction + 1
    End If
    If strTestCycle <> "N/A" Then
        iAction = iAction + 2
    End If

    ' Create the dictionary object
    Set objTestFunctionalAreaDictionary = New Dictionary

    '   Set up the test set filter
    Set tdcTestSetFactory = tdc.TestSetFactory
    Set tdcTestFilter = tdcTestSetFactory.Filter
    tdcTestFilter.Filter(strProjectCycleLabel) = Chr(34) & strProjectName & Chr(34)
    tdcTestFilter.Filter(strTestPhaseCycleLabel) = Chr(34) & strTestPhase & Chr(34)
    Select Case iAction
        Case 1
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
        Case 2
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
        Case 3
            tdcTestFilter.Filter(strSubProjectCycleLabel) = Chr(34) & strSubProjectName & Chr(34)
            tdcTestFilter.Filter(strTestCycleLabel) = Chr(34) & strTestCycle & Chr(34)
    End Select
    
    Set colTestSets = tdcTestSetFactory.NewList(tdcTestFilter.Text)

    For Each objTestSet In colTestSets
        rc = objTestSet.TestSetFolder.Path
        If InStr(1, rc, "Unattached") = 0 And InStr(1, rc, "99. Archive") = 0 And InStr(1, rc, "Trash") = 0 Then
            '   Get the count by this filter
            ReturnFunctionalAreaCount objTestSet, "No Run"
            ReturnFunctionalAreaCount objTestSet, "Not Completed"
            ReturnFunctionalAreaCount objTestSet, "Failed"
            ReturnFunctionalAreaCount objTestSet, "Passed"
        End If
    Next
    
    '   Sort the array
    SortDictionary objTestFunctionalAreaDictionary, True
    
    '   Return the data
    CountTestsByFunctionalArea = objTestFunctionalAreaDictionary.Keys
    
    '   Release objects
    Set tdcTestFilter = Nothing
    Set colTestSets = Nothing
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function ReturnFunctionalAreaCount(ByRef objTestSet, ByVal strStatus As String)
Dim TSTestFact
Dim tdcTestFilter

        '   Set the TestSet Factory object
        Set TSTestFact = objTestSet.TSTestFactory
        '   Set up the filter
        Set tdcTestFilter = TSTestFact.Filter
        tdcTestFilter.Filter("TC_STATUS") = Chr(34) & strStatus & Chr(34)
                
        '   Get the list from the filter
        Set colTSTests = tdcTestFilter.NewList
        
        '   Loop round the tests
        For Each objTest In colTSTests
            strThisFunctionalArea = objTest.Field(strFunctionalAreaTestLabel)
            If strThisFunctionalArea = "" Then
                strThisFunctionalArea = "Not Assigned"
            End If
            '   See if this functional area exists in the original list
            For i = 0 To objTestFunctionalAreaDictionary.Count - 1
                If InStr(1, objTestFunctionalAreaDictionary.Keys(i), strThisFunctionalArea) Then
                    blnOK = True
                    Exit For
                End If
            Next
            '   See if we've not got the item in our list
            If blnOK = False Then
                '   Add the new item to the dictionary with its severities
                objTestFunctionalAreaDictionary.Add "Passed" & "|" & strThisFunctionalArea, 0
                objTestFunctionalAreaDictionary.Add "Failed" & "|" & strThisFunctionalArea, 0
                objTestFunctionalAreaDictionary.Add "Not Completed" & "|" & strThisFunctionalArea, 0
                objTestFunctionalAreaDictionary.Add "No Run" & "|" & strThisFunctionalArea, 0
            End If
            intCount = CInt(objTestFunctionalAreaDictionary.Item(strStatus & "|" & strThisFunctionalArea))
            objTestFunctionalAreaDictionary.Item(strStatus & "|" & strThisFunctionalArea) = intCount + 1
            blnOK = False
        Next
        
        '   Release objects
        Set TSTestFact = Nothing
        Set tdcTestFilter = Nothing
        Set colTSTests = Nothing
        
End Function


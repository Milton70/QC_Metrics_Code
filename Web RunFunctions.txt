Attribute VB_Name = "RunFunctions"
Public Function GetRuns()
Dim blnErrorPage As Boolean
iTI = 0
iPass = 0
iFail = 0
iNC = 0
iTR = 0
iDum = 0
ReDim Preserve arrTestInstance(iTI)
arrTestInstance(iTI) = "Project Start|0"
ReDim Preserve arrPassed(iPass)
arrPassed(iPass) = "Project Start|0"
ReDim Preserve arrFailed(iFail)
arrFailed(iFail) = "Project Start|0"
ReDim Preserve arrNC(iNC)
arrNC(iNC) = "Project Start|0"
ReDim Preserve arrTR(iTR)
arrTR(iTR) = "Project Start|0"
ReDim Preserve arrDummy(iDum)
arrDummy(iDum) = "Project Start|0"

    myDate = myStartDate
    
    '   Break out of this function if we've got no start date
    If myDate = "00:00:00" Then
        blnErrorPage = True
        GoTo WriteOutHTML
    End If
    
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
        myDate = RunsMonthsWeeks(arrMonths, myDate)
    End If
    '   Write out the week values
    If arrWeekends(0) <> "No Weeks" Then
        myDate = RunsMonthsWeeks(arrWeekends, myDate)
    End If

    '   See if myDate already exists in the array (it'll be the last one)
    If InStr(1, arrTestInstance(UBound(arrTestInstance)), myDate) > 0 Then
        myDate = myDate + 1
    End If
    
    '   Do the daily values upto todays date
    Do While myDate <= TodaysDate
        
        strDateFilter = myDate
        
        '   Get value from run module for this date
        myValue = CountRunsByDate(strDateFilter)
        
        '   Split the value
        mySplit = Split(myValue, "|")
        
        '   Add to the arrays
        iTI = iTI + 1
        ReDim Preserve arrTestInstance(iTI)
        arrTestInstance(iTI) = strDateFilter & "|" & mySplit(0)
        iTR = iTR + 1
        ReDim Preserve arrTR(iTR)
        arrTR(iTR) = strDateFilter & "|" & mySplit(1)
        iPass = iPass + 1
        ReDim Preserve arrPassed(iPass)
        arrPassed(iPass) = strDateFilter & "|" & mySplit(2)
        iFail = iFail + 1
        ReDim Preserve arrFailed(iFail)
        arrFailed(iFail) = strDateFilter & "|" & mySplit(3)
        iNC = iNC + 1
        ReDim Preserve arrNC(iNC)
        arrNC(iNC) = strDateFilter & "|" & mySplit(4)
        iDum = iDum + 1
        ReDim Preserve arrDummy(iDum)
        arrDummy(iDum) = strDateFilter & "|0"
        
        '   Make these cumulative
        If iTI > 0 Then
            myoldsplit = Split(arrTestInstance(iTI - 1), "|")
            mycurrsplit = Split(arrTestInstance(iTI), "|")
            iTITot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrTestInstance(iTI) = mycurrsplit(0) & "|" & iTITot
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
        If iTR > 0 Then
            myoldsplit = Split(arrTR(iTR - 1), "|")
            mycurrsplit = Split(arrTR(iTR), "|")
            iTRTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrTR(iTR) = mycurrsplit(0) & "|" & iTRTot
        End If
ExitLoop:
        myDate = myDate + 1
    Loop
    
    '   Now get just the planned values out into the future if required
    If dateEndDate > TodaysDate Then
        iDays = DateDiff("d", TodaysDate, dateEndDate)
        If iDays <= 30 Then
            myDate = DateAdd("d", 1, TodaysDate)
            Do While myDate <= dateEndDate
                iDum = iDum + 1
                ReDim Preserve arrDummy(iDum)
                arrDummy(iDum) = myDate & "|0"
                myDate = myDate + 1
            Loop
            GoTo WriteOutHTML
        End If
        myDate = DateAdd("d", 1, TodaysDate)
        '   Get the date next week
        dateNextWeek = DateAdd("d", 7, TodaysDate)
        '   See if it's past our end date
        If dateNextWeek >= dateEndDate Then
            '   Get the last value reported cos we're just going to repeat this
            Do While myDate <= dateEndDate
                iDum = iDum + 1
                ReDim Preserve arrDummy(iDum)
                arrDummy(iDum) = myDate & "|0"
                myDate = myDate + 1
            Loop
            GoTo WriteOutHTML
        End If
        dateNextMonth = DateAdd("m", 1, dateNextWeek)
        '   See if it's past our end date
        If dateNextMonth >= dateEndDate Then
            arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
            If arrWeekends(0) <> "No Weeks" Then
                If arrWeekends(0) < dateEndDate Then
                    Do While myDate <= dateNextWeek
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = myDate & "|0"
                        myDate = myDate + 1
                    Loop
                    For Each Ele In arrWeekends
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = Ele & "|0"
                        myDate = Ele
                    Next
                    myDate = myDate + 1
                    Do While myDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = myDate & "|0"
                        myDate = myDate + 1
                    Loop
                Else
                    Do While myDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = myDate & "|0"
                        myDate = myDate + 1
                    Loop
                End If
            Else
                Do While myDate <= dateEndDate
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = myDate & "|0"
                    myDate = myDate + 1
                Loop
                GoTo WriteOutHTML
            End If
        End If
        '   See how many days we're dealing with
        iDays = DateDiff("d", dateNextMonth, dateEndDate)
            If iDays < 0 Then
                GoTo WriteOutHTML
            End If
            If iDays <= 30 Then
                arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
                '   Write out the first weeks worth of days
                Do While myDate <= dateNextWeek
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = myDate & "|0"
                    myDate = myDate + 1
                Loop
                '   Now write out the remaining weeks
                For Each Ele In arrWeekends
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = Ele & "|0"
                    myDate = Ele
                Next
                '   See if we need to write out the final few days
                If myDate < dateEndDate Then
                    myDate = myDate + 1
                    Do While myDate <= dateEndDate
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = myDate & "|0"
                    myDate = myDate + 1
                Loop
                End If
            Else
                arrMonths = FindMonths(dateNextMonth, dateEndDate)
                arrWeekends = FindWeekends(dateNextWeek, dateNextMonth)
                Do While myDate <= dateNextWeek
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = myDate & "|0"
                    myDate = myDate + 1
                Loop
                For Each Ele In arrWeekends
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = Ele & "|0"
                    myDate = Ele
                Next
                For Each Ele In arrMonths
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = Ele & "|0"
                    myDate = Ele
                Next
            End If
        
    End If
    
WriteOutHTML:

    '   See if we're defaulting to the error page or not
    If blnErrorPage = False Then
    
        '   Now write out the Detected vs Closed table info
        fso.CopyFile strTemplatePath & "TestRunsTableTemplate.txt", strFolderPath & strPathandFileName & "-TestRunsTable.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestRunsTable.asp", ForAppending, True)
        
        '   Write out the table header details etc
        myFile.WriteLine "<table border=1 align=center>"
        myFile.WriteLine "<tr><th colspan=7 align=center>" & strHeader & "</th></tr>"
        myFile.WriteLine "<tr bgcolor='#56A5EC'><td>&nbsp;</td><td>Total Test Instance</td><td>Total Test Runs</td><td>Test Runs Passed</td><td>Test Runs Failed</td><td>Test Runs N/C</td></tr>"
        iCount = UBound(arrTestInstance)
        iDummyCount = UBound(arrDummy)
        If iCount = iDummyCount Then
            iTotal = iDummyCount
        Else
            iTotal = iCount
        End If
        i = 0
        Do
            '   Get the values from each of the arrays for this array element
            aSplit = Split(arrTestInstance(i), "|")
            bSplit = Split(arrTR(i), "|")
            cSplit = Split(arrPassed(i), "|")
            dSplit = Split(arrFailed(i), "|")
            eSplit = Split(arrNC(i), "|")
            fSplit = Split(arrDummy(i), "|")
            
            '   If we're on the first element then just default to project start and zeros
            If i = 0 Then
                myFile.WriteLine "<tr bgcolor ='#B4CFEC'><td>Project Start</td><td>0</td><td>0</td><td>0</td><td>0</td><td>0</td></tr>"
            Else
                If Weekday(aSplit(0)) <> 7 And Weekday(aSplit(0)) <> 1 Then
                    If IsEven(i) = True Then
                        myFile.WriteLine ("<tr bgcolor ='#B4CFEC'>")
                    Else
                        myFile.WriteLine ("<tr bgcolor ='ivory'>")
                    End If
                    '   Write a row to the table
                    myFile.WriteLine "<td>" & aSplit(0) & "</td><td>" & aSplit(1) & "</td><td>" & bSplit(1) & "</td><td>" & cSplit(1) & "</td><td>" & dSplit(1) & "</td><td>" & eSplit(1) & "</td></tr>"
                End If
            End If
            i = i + 1
        Loop Until i > iTotal
        If iDummyCount > iCount Then
            For j = i To UBound(arrDummy)
                fSplit = Split(arrDummy(j), "|")
                If fSplit(0) = "Project Start" Then
                    myFile.WriteLine "<td>" & fSplit(0) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
                Else
                    If Weekday(fSplit(0)) <> 7 And Weekday(fSplit(0)) <> 1 Then
                        If IsEven(i) = True Then
                            myFile.WriteLine ("<tr bgcolor ='#B4CFEC'>")
                        Else
                            myFile.WriteLine ("<tr bgcolor ='ivory'>")
                        End If
                        '   Write a row to the table
                        myFile.WriteLine "<td>" & fSplit(0) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
                    End If
                End If
            Next
        End If
        '   Write the remainder of the html
        myFile.WriteLine "</table></body></html>"
        '   Close the file
        myFile.Close
    
        '   Now write out the test run info
        Set mySource = fso.OpenTextFile(strTemplatePath & "TestRunsTemplate.txt", ForReading)
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestRuns.aspx", ForWriting, True)
        Do While mySource.AtEndOfStream <> True
            rc = mySource.ReadLine
            If InStr(1, rc, "TestInstancePoints") > 0 Then
                '   Loop round our Baseline array
                For i = 0 To UBound(arrTestInstance)
                    '   Split the file
                    mySplit = Split(arrTestInstance(i), "|")
                    If mySplit(0) = "Project Start" Then
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                    Else
                        If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                            '   Re-format the date part
                            myDate = Format(mySplit(0), "dd mmm yy")
                            '   Write the value
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                        End If
                    End If
                Next
            Else
                If InStr(1, rc, "PassedPoints") > 0 Then
                    '   Loop round our Replanned array
                    For i = 0 To UBound(arrPassed)
                        '   Split the file
                        mySplit = Split(arrPassed(i), "|")
                        If mySplit(0) = "Project Start" Then
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                        Else
                            If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                '   Re-format the date part
                                myDate = Format(mySplit(0), "dd mmm yy")
                                '   Write the value
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                            End If
                        End If
                    Next
                Else
                    If InStr(1, rc, "FailedPoints") > 0 Then
                        '   Loop round our Executed array
                        For i = 0 To UBound(arrFailed)
                            '   Split the file
                            mySplit = Split(arrFailed(i), "|")
                            If mySplit(0) = "Project Start" Then
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                            Else
                                If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                    '   Re-format the date part
                                    myDate = Format(mySplit(0), "dd mmm yy")
                                    '   Write the value
                                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                                End If
                            End If
                        Next
                    Else
                        If InStr(1, rc, "NCPoints") > 0 Then
                            '   Loop round our passed array
                            For i = 0 To UBound(arrNC)
                                '   Split the file
                                mySplit = Split(arrNC(i), "|")
                                If mySplit(0) = "Project Start" Then
                                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                                Else
                                    If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                        '   Re-format the date part
                                        myDate = Format(mySplit(0), "dd mmm yy")
                                        '   Write the value
                                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                                    End If
                                End If
                            Next
                        Else
                            If InStr(1, rc, "TotalPoints") > 0 Then
                                '   Loop round our failed array
                                For i = 0 To UBound(arrTR)
                                    '   Split the file
                                    mySplit = Split(arrTR(i), "|")
                                    If mySplit(0) = "Project Start" Then
                                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                                    Else
                                        If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                            '   Re-format the date part
                                            myDate = Format(mySplit(0), "dd mmm yy")
                                            '   Write the value
                                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                                        End If
                                    End If
                                Next
                            Else
                                If InStr(1, rc, "HiddenPoints") > 0 Then
                                    '   Loop round our failed array
                                    For i = 0 To UBound(arrDummy)
                                        '   Split the file
                                        mySplit = Split(arrDummy(i), "|")
                                        If mySplit(0) = "Project Start" Then
                                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                                        Else
                                            If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                                '   Re-format the date part
                                                myDate = Format(mySplit(0), "dd mmm yy")
                                                '   Write the value
                                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                                            End If
                                        End If
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
        
        '   Open the supporting test set page and include the graph just created
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        '   Replace the graph
        strText = Replace(strText, "TestsRuns", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestRuns.aspx" & Chr(34))
        
        '   Write it back out
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForWriting, True)
        myFile.WriteLine strText
        myFile.Close
    Else
        '   Copy the defect error page to the test runs page
        fso.CopyFile strTemplatePath & "TestSetsMissingTemplate.txt", strFolderPath & strPathandFileName & "-TestSetDetails.asp"
        '   Open the file and change the header
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        '   Replace the text
        strText = Replace(strText, "strHeader", "No Test Run Details for " & strHeader)
        
        '   Write it back out
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForWriting, True)
        myFile.WriteLine strText
        myFile.Close
    
    End If
End Function
Public Function CountRunsByDate(ByVal strDateFilter As String)
Dim tdcTestSetFactory
Dim tdcTestFilter
Dim intTestCount As Integer
Dim colTestSets As list
Dim iAction As Integer
Dim myArr()
Dim blnTestSetsFound As Boolean
iAction = 0
iTestInstanceCount = 0
iTestRunsCount = 0
iRunPassedCount = 0
iRunFailedCount = 0
iRunNCCount = 0

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
    
    i = -1

    For Each objTestSet In colTestSets
        rc = objTestSet.TestSetFolder.Path
        If InStr(1, rc, "Unattached") = 0 And InStr(1, rc, "99. Archive") = 0 And InStr(1, rc, "Trash") = 0 Then
                i = i + 1
                ReDim Preserve myArr(i)
                myArr(i) = ReturnRunCount(objTestSet, strDateFilter)
                blnTestSetsFound = True
        End If
    Next
    
    '   Split the array and total up for the date
    If blnTestSetsFound = True Then
        For Each Ele In myArr
            mySplit = Split(Ele, "|")
            iTestInstanceCount = iTestInstanceCount + mySplit(0)
            iTestRunsCount = iTestRunsCount + mySplit(1)
            iRunPassedCount = iRunPassedCount + mySplit(2)
            iRunFailedCount = iRunFailedCount + mySplit(3)
            iRunNCCount = iRunNCCount + mySplit(4)
        Next
    Else
        iTestInstanceCount = 0
        iTestRunsCount = 0
        iRunPassedCount = 0
        iRunFailedCount = 0
        iRunNCCount = 0
    End If
    
    '   Return values
    CountRunsByDate = iTestInstanceCount & "|" & iTestRunsCount & "|" & iRunPassedCount & "|" & iRunFailedCount & "|" & iRunNCCount

    '   Clean up objects
    Set tdcTestFilter = Nothing
    Set colTestSets = Nothing
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function ReturnRunCount(ByRef objTestSet, ByVal strDateFilter As String)
Dim objTSTestFactory
Dim tdcTestFilter
Dim iTestInstanceCount As Integer
Dim iTestRunsCount As Integer
Dim iRunPassedCount As Integer
Dim iRunFailedCount As Integer
Dim iRunNCCount As Integer
iTestInstanceCount = 0
iTestRunsCount = 0
iRunPassedCount = 0
iRunFailedCount = 0
iRunNCCount = 0

    '   Set the TestSet Factory object
    Set objTSTestFactory = objTestSet.TSTestFactory
        '   Set up the filter
        Set tdcTestFilter = objTSTestFactory.Filter
        Set colTSTests = tdcTestFilter.NewList
        For Each TSTest In colTSTests
            
            Set tdcTSRunFactory = TSTest.RunFactory
            Set tdcTSRunFilter = tdcTSRunFactory.Filter
            tdcTSRunFilter.Filter("RN_EXECUTION_DATE") = strDateFilter
            Set tdcRuns = tdcTSRunFactory.NewList(tdcTSRunFilter.Text)

            '   Loop round runs
            For Each TestRun In tdcRuns
                '   Add to the test runs count
                iTestRunsCount = iTestRunsCount + 1
                strStatus = TestRun.Status
                Select Case strStatus
                    Case "Passed"
                        iRunPassedCount = iRunPassedCount + 1
                    Case "Failed"
                        iRunFailedCount = iRunFailedCount + 1
                    Case "Not Completed"
                        iRunNCCount = iRunNCCount + 1
                End Select
            Next
        Next
        
        '   Re-set up the filter
        Set tdcTestFilter = objTSTestFactory.Filter
        tdcTestFilter.Filter("TC_EXEC_DATE") = strDateFilter
        Set colTSTests = tdcTestFilter.NewList
        iTestInstanceCount = colTSTests.Count
        
        '   Return the count
        ReturnRunCount = iTestInstanceCount & "|" & iTestRunsCount & "|" & iRunPassedCount & "|" & iRunFailedCount & "|" & iRunNCCount

        '   Clean up objects
        Set objTSTestFactory = Nothing
        Set tdcTestFilter = Nothing
        Set colTSTests = Nothing


End Function
Public Function RunsMonthsWeeks(ByVal arrArray As Variant, ByVal dateThisStartDate As Date)
    
    If arrArray(0) <> "No Month" Then
    
        For Each Ele In arrArray
            
            strDateFilter = ">= " & dateThisStartDate & " And < " & Ele
            
            myValues = CountRunsByDate(strDateFilter)
            mySplit = Split(myValues, "|")
                
            iTI = iTI + 1
            ReDim Preserve arrTestInstance(iTI)
            arrTestInstance(iTI) = Ele & "|" & mySplit(0)
            iTR = iTR + 1
            ReDim Preserve arrTR(iTR)
            arrTR(iTR) = Ele & "|" & mySplit(1)
            iPass = iPass + 1
            ReDim Preserve arrPassed(iPass)
            arrPassed(iPass) = Ele & "|" & mySplit(2)
            iFail = iFail + 1
            ReDim Preserve arrFailed(iFail)
            arrFailed(iFail) = Ele & "|" & mySplit(3)
            iNC = iNC + 1
            ReDim Preserve arrNC(iNC)
            arrNC(iNC) = Ele & "|" & mySplit(4)
            iDum = iDum + 1
            ReDim Preserve arrDummy(iDum)
            arrDummy(iDum) = Ele & "|0"
            
            '   Make these cumulative
            If iTI > 0 Then
                myoldsplit = Split(arrTestInstance(iTI - 1), "|")
                mycurrsplit = Split(arrTestInstance(iTI), "|")
                iTITot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrTestInstance(iTI) = mycurrsplit(0) & "|" & iTITot
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
            If iTR > 0 Then
                myoldsplit = Split(arrTR(iTR - 1), "|")
                mycurrsplit = Split(arrTR(iTR), "|")
                iTRTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrTR(iTR) = mycurrsplit(0) & "|" & iTRTot
            End If
              
            dateThisStartDate = Ele
            
        Next
    End If
    
    '   Return the date for the next section
    RunsMonthsWeeks = dateThisStartDate
    
End Function
Public Function RunAutomationReport()
Dim tdcTestSetFactory, Runfact, myStepFact
Dim tdcTestFilter
Dim colTestSets
Dim TSTestFact
Dim objTestSet, objTest, objStep
Dim colTSTests, colBPRuns, colStep
Dim myStepFilter
Dim myArr()
Dim iCount As Long
Dim x As Integer
Dim y As Integer
Dim blnWriteHeader As Boolean
x = 1
    '   Copy the template
    fso.CopyFile strTemplatePath & "AutomationReportTemplate.txt", strFolderPath & strPathandFileName & "-AutomationReport.asp"
    '   Open the file for writing
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-AutomationReport.asp", ForAppending, True)

    '   See which kind of search we're doing
    If strSubProjectName <> "N/A" Then
        iAction = iAction + 1
    End If
    If strTestCycle <> "N/A" Then
        iAction = iAction + 2
    End If
    
    '   Set up test set filter
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
    
    '   Get the returned list
    Set colTestSets = tdcTestSetFactory.NewList(tdcTestFilter.Text)
    
    '   See if we've got any test sets
    If colTestSets.Count = 0 Then
        myFile.WriteLine "<table align='center' border=0 width='90%'>"
        myFile.WriteLine "<tr><th>strHeader</th></tr>"
        myFile.WriteLine "<tr></tr>"
        myFile.WriteLine "<tr valign=top>"
        myFile.WriteLine "<td>"
        myFile.WriteLine "<table align=center border='0'>"
        myFile.WriteLine "<tr bgcolor='#CCCCCC'><td>Sorry but No Automation Test Sets have been found.</td></tr>"
        myFile.WriteLine "</table>"
        myFile.WriteLine "</td></tr></table>"
        myFile.Close
        Exit Function
    End If

    '   Loop round test sets
    For Each objTestSet In colTestSets
        
        '   Set the TestSet Factory object
        Set TSTestFact = objTestSet.TSTestFactory
        '   Set up the filter
        Set tdcTestFilter = TSTestFact.Filter
        
        '   Set the filter value to failed status
        tdcTestFilter.Filter("TC_STATUS") = Chr(34) & "Failed" & Chr(34)
        
        '   Get the list from the filter
        Set colTSTests = tdcTestFilter.NewList
        
        '   See if we've got any or not
        If colTSTests.Count = 0 Then
            myFile.WriteLine "<table align='center' border=0 width='90%'>"
            myFile.WriteLine "<tr><th>strHeader</th></tr>"
            myFile.WriteLine "<tr></tr>"
            myFile.WriteLine "<tr valign=top>"
            myFile.WriteLine "<td>"
            myFile.WriteLine "<table align=center border='0'>"
            myFile.WriteLine "<tr bgcolor='#CCCCCC'><td>All Automation Test Sets have Passed.</td></tr>"
            myFile.WriteLine "</table>"
            myFile.WriteLine "</td></tr></table>"
            myFile.Close
            Exit Function
        End If
        
        '   Write out the table header section
        myFile.WriteLine "<table style='table-layout:fixed;font-size:10px;border=1;word-wrap:break-word;cellspacing=1;text-align:center;width=100%;background-color:ivory'>"
        If blnWriteHeader = False Then
            myFile.WriteLine "<thead><tr><th colspan=7>strHeader</th></tr><tr></tr>"
            blnWriteHeader = True
        Else
            myFile.WriteLine "<thead><tr><th colspan=7>&nbsp;</th></tr><tr></tr>"
        End If
        myFile.WriteLine "<tr align=center>"
        myFile.WriteLine "<th bgcolor='#99CC99' width=10%>Test Set</th><th bgcolor='#99CC99' width=15%>Test Name</th><th bgcolor='#99CC99' width=15%>Status</th><th bgcolor='#99CC99' width=10%>Exec Date</th><th bgcolor='#99CC99' width=15%>Run Name</th><th bgcolor='#99CC99'width=10%>Host</th><th bgcolor='#99CC99'width=10%>Environment</th></tr></tr></thead>"
        
        '   Loop round the tests returned from the filter
        For Each objTest In colTSTests
            
            '   Get the run factory from the test
            Set Runfact = objTest.RunFactory
            Set colBPRuns = Runfact.NewList("")
            
            '   See if the last run is a failure
            If colBPRuns.Item(colBPRuns.Count).Status = "Failed" Then
                y = 1
                
                '   Write the main details
                myFile.WriteLine "<tr align=center>"
                If IsEven(x) = False Then
                    myFile.WriteLine "<td bgcolor='#ECFDEC'><a href=testdirector:cmutility:8080/qcbin,LCHC_STREAM," & strQCProject & ",;7:" & objTestSet.ID & ">" & objTestSet.Name & "</a></td>"
                    myFile.WriteLine "<td bgcolor='#ECFDEC'>" & objTest.Name & "</td>"
                    myFile.WriteLine "<td bgcolor='#ECFDEC' onclick='toggle(" & Chr(34) & "Failed_" & iFails & Chr(34) & ");' style='cursor:pointer'><u><font color='red'>" & objTest.Status & "</font></u></td>"
                    myFile.WriteLine "<td bgcolor='#ECFDEC'>" & objTest.Field("TC_EXEC_DATE") & "</td>"
                    myFile.WriteLine "<td bgcolor='#ECFDEC'>" & colBPRuns.Item(colBPRuns.Count).Name & "</td>"
                    myFile.WriteLine "<td bgcolor='#ECFDEC'>" & colBPRuns.Item(colBPRuns.Count).Field("RN_HOST") & "</td>"
                    myFile.WriteLine "<td bgcolor='#ECFDEC'>" & objTest.Field(GetFieldName("Test Environment", "TESTCYCL")) & "</td>"
                Else
                    myFile.WriteLine "<td><a href=testdirector:cmutility:8080/qcbin,LCHC_STREAM," & strQCProject & ",;7:" & objTestSet.ID & ">" & objTestSet.Name & "</a></td>"
                    myFile.WriteLine "<td>" & objTest.Name & "</td>"
                    myFile.WriteLine "<td onclick='toggle(" & Chr(34) & "Failed_" & iFails & Chr(34) & ");' style='cursor:pointer'><u><font color='red'>" & objTest.Status & "</font></u></td>"
                    myFile.WriteLine "<td>" & objTest.Field("TC_EXEC_DATE") & "</td>"
                    myFile.WriteLine "<td>" & colBPRuns.Item(colBPRuns.Count).Name & "</td>"
                    myFile.WriteLine "<td>" & colBPRuns.Item(colBPRuns.Count).Field("RN_HOST") & "</td>"
                    myFile.WriteLine "<td>" & objTest.Field(GetFieldName("Test Environment", "TESTCYCL")) & "</td>"
                End If
                ' up x
                x = x + 1
                myFile.WriteLine "</tr>"
                myFile.WriteLine "<tbody id='Failed_" & iFails & "' style='display:none' align='center'>"
                myFile.WriteLine "<tr bgcolor='#99CC99'>"
    
                ' Now the Steps
                Set myStepFact = colBPRuns.Item(colBPRuns.Count).StepFactory
                Set myStepFilter = myStepFact.Filter
                myStepFilter.Filter("ST_STATUS") = Chr(34) & "Failed" & Chr(34)
                Set colStep = myStepFact.NewList(myStepFilter.Text)
                '   Stick em in an array
                Erase myArr
                iCount = -1
                For Each objStep In colStep
                    iCount = iCount + 1
                    ReDim Preserve myArr(iCount)
                    myArr(iCount) = objStep.Name & "|" & Replace(objStep.Field("ST_DESCRIPTION"), vbCrLf, " ")
                Next
                '   Loop round the array so we can look forward
                For j = LBound(myArr) To UBound(myArr)
                    
                    '   Split the string
                    mySplit = Split(myArr(j), "|")
                    '   See if we've got an iteration
                    If InStr(1, mySplit(0), "Test Iteration") > 0 Then
                        myFile.WriteLine "<tr><th>&nbsp;</th><th bgcolor='#99CC99'>" & mySplit(0) & "</th></tr>"
                    Else
                        
                        '   See if the next item is a Component
                        If mySplit(1) = "" Then
                            myFile.WriteLine "<tr><th colspan=2>&nbsp;</th><th bgcolor='#99CC99'>Component Name</th><th bgcolor='#99CC99'>Status</th></tr>"
                            myFile.WriteLine "<tr><td colspan=2>&nbsp;</td><td>" & mySplit(0) & "</td><td><font color='red'>Failed</font></td></tr>"
                        Else
                            myFile.WriteLine "<tr><th colspan=3>&nbsp;</th><th bgcolor='#99CC99'>Step Name</th><th bgcolor='#99CC99'>Step Status</th><th bgcolor='#99CC99' colspan=2>Step Description</th></tr>"
                            If IsEven(y) = False Then
                                myFile.WriteLine "<tr><td colspan=3>&nbsp;</td><td bgcolor='#ECFDEC'>" & mySplit(0) & "</td><td bgcolor='#ECFDEC'><font color='red'>Failed</font></td><td colspan=2 bgcolor='#ECFDEC'>" & mySplit(1) & "</td></tr>"
                            Else
                                myFile.WriteLine "<tr><td colspan=3>&nbsp;</td><td>" & mySplit(0) & "</td><td><font color='red'>Failed</font></td><td colspan=2>" & mySplit(1) & "</td></tr>"
                            End If
                            '    up y
                            y = y + 1
                        End If
                    End If
                Next
                '   End the tbody
                myFile.WriteLine "</tr></tbody>"
                iFails = iFails + 1
            End If
        Next
    Next
    
    '   Finish off html and close file
    myFile.WriteLine "</table><%FinishPage();%></body></html>"
    myFile.Close
    
    '   Clean up objects
    Set tdcTestSetFactory = Nothing
    Set tdcTestFilter = Nothing
    Set colTestSets = Nothing
    
    
End Function

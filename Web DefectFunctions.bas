Attribute VB_Name = "DefectFunctions"
Public Function GetOpenDefects()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
Dim strStatus, strPriority, strSeverity, strSummary
Dim strDetectedOnDate, strAssignedTo, strDefectId
Dim iRow
Dim arrReturn()
Dim blnNoBugs As Boolean

        '   Set up the project and test phase filter
        Set tdcBugFactory = tdc.BugFactory
        Set tdcBugFilter = tdcBugFactory.Filter
        tdcBugFilter.Filter("BG_PROJECT") = Chr(34) & strProjectName & Chr(34)
        tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(34) & strTestPhase & Chr(34)
        '   See if we're filtering on Sub-Project
        If strSubProjectName <> "N/A" Then
            tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
        End If
        tdcBugFilter.Filter("BG_STATUS") = "(New Or Assigned Or Open Or Reopen Or " & Chr(34) & "Failed Testing" & Chr(34) & ")"
        tdcBugFilter.Filter("BG_DETECTION_DATE") = "<= " & TodaysDate
   
        '  Get the list of bugs
        Set colBugList = tdcBugFilter.NewList

        '  Set j to be the array counter
        j = 0
    
        '   See if we've got any bugs
        If colBugList.Count = 0 Then
            blnNoBugs = True
        End If
        
        '   Get the details by severity
        '   1-Critical
        For i = 1 To colBugList.Count
            '  Get the severity
            strSeverity = colBugList.Item(i).Field("BG_SEVERITY")

            '  See if we've got critical bugs
            If strSeverity = "1-Critical" Then
                    '   Get the defect details
                    strDefectId = colBugList.Item(i).ID
                    strSummary = colBugList.Item(i).Summary
                    strDetectedOnDate = colBugList.Item(i).Field(strDetectedOnDateLabel)
                    strStatus = colBugList.Item(i).Status
                    strAssignedTo = colBugList.Item(i).AssignedTo
                    strPriority = colBugList.Item(i).Priority
            
                    '   Write out the defect details to array
                    ReDim Preserve arrReturn(j)
                    arrReturn(j) = strDefectId & "|" & strSummary & "|" & strDetectedOnDate & "|" & strStatus & "|" & strAssignedTo & "|" & strSeverity & "|" & strPriority
            
                    '   Up the counter
                    j = j + 1
            End If
        Next
        '   2-High
        For i = 1 To colBugList.Count
            '  Get the severity
            strSeverity = colBugList.Item(i).Field("BG_SEVERITY")
        
            If strSeverity = "2-High" Then
        
                    '   Get the defect details
                    strDefectId = colBugList.Item(i).ID
                    strSummary = colBugList.Item(i).Summary
                    strDetectedOnDate = colBugList.Item(i).Field(strDetectedOnDateLabel)
                    strStatus = colBugList.Item(i).Status
                    strAssignedTo = colBugList.Item(i).AssignedTo
                    strPriority = colBugList.Item(i).Priority
            
                    '   Write out the defect details to array
                    ReDim Preserve arrReturn(j)
                    arrReturn(j) = strDefectId & "|" & strSummary & "|" & strDetectedOnDate & "|" & strStatus & "|" & strAssignedTo & "|" & strSeverity & "|" & strPriority
            
                    '   Up the row
                    j = j + 1
            End If
    Next
    
    '  Open the Open Defects template
    Set myFile = fso.OpenTextFile(strTemplatePath & "OpenDefectsTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    '   Replace the title data
    strNewText = Replace(strText, "ProjectTitle", strPathandFileName)
    
    '   Write the new text to the asp file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-OpenDefects.asp", ForWriting, True)
    myFile.WriteLine strNewText
    myFile.Close
    '   Open the file again, for appending to
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-OpenDefects.asp", ForAppending, True)
    
    myFile.WriteLine ("<body>")
    myFile.WriteLine ("<table width='700' align='center' border='1'>")
    myFile.WriteLine ("<tr><th colspan=7>" & strHeader & "</th></tr>")
    myFile.WriteLine ("<tr><th center colspan='7' bgcolor='#99CC66'>Defect Details - Critical / High Severity - New /Assigned / Open / Reopen / Failed Testing</th></tr>")
    myFile.WriteLine ("<tr bgcolor='#CCCCCC'><th><font size='-1'>Defect ID</font></th><th><font size='-1'>Summary</font></th><th><font size='-1'>Detected On Date</font></th>")
    myFile.WriteLine ("<th><font size='-1'>Status</font></th><th><font size='-1'>Assigned To</font></th><th><font size='-1'>Severity</font></th><th><font size='-1'>Priority</font></th></tr>")
    
    '   See if we've got any bugs at all
    If blnNoBugs = False Then
        '   See if we've got an array or not
        If j > 0 Then
            '  Loop round the array putting values into the correct columns
            For i = 0 To UBound(arrReturn)
                '  Split the array element
                mySplit = Split(arrReturn(i), "|")
                '  See if we colour this row
                If IsEven(i) = True Then
                    myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
                Else
                    myFile.WriteLine ("<tr bgcolor ='ivory'>")
                End If
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(0) + "</font></td>")
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(1) + "</font></td>")
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(2) + "</font></td>")
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(3) + "</font></td>")
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(4) + "</font></td>")
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(5) + "</font></td>")
                myFile.WriteLine ("<td><font size='-1'>" + mySplit(6) + "</font></td>")
                myFile.WriteLine ("</tr>")
            Next
        End If
    End If
    myFile.WriteLine ("</table>")
    myFile.WriteLine ("<%FinishPage();%>")
    myFile.WriteLine ("</body>")
    myFile.WriteLine ("</html>")
    myFile.Close
    
    
End Function
Public Function GetDefectsBySeverityByPriority()

    '   First see if we've got any bugs
    rc = GetDefectCount(False)
    If rc = False Then
        '   Copy the defect error page to the test runs page
        fso.CopyFile strTemplatePath & "DefectsMissingTemplate.txt", strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt"
        '   Open the file and change the header
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        '   Replace the text
        strText = Replace(strText, "strHeader", "No Defect Details for " & strHeader)
        
        '   Write it back out
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForWriting, True)
        myFile.WriteLine strText
        myFile.Close
        
        '   Set the flag
        blnDontRun = True
        Exit Function
    End If

    '   Get the defect arrays
    myPriority = DefectStatusByPriority()
    mySeverity = DefectStatusBySeverity()
    
    '  Open the template file
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectsStatusTemplate.txt", ForReading)
    '   Get the text
    strText = myFile.ReadAll
    '   Close this file
    myFile.Close
    
    '   Replace the header
    strNewText = Replace(strText, "strHeader", strHeader)
    
    '  Write out the html
    iS1 = 0
    iS2 = 0
    iS3 = 0
    iS4 = 0
    '  Write out the first four status
    For iSev = 0 To 4
        
        '  Split the record
        mySplit = Split(mySeverity(iSev), "|")
        '  Keep running total of each status
        iS1 = iS1 + CInt(mySplit(1))
        iS2 = iS2 + CInt(mySplit(2))
        iS3 = iS3 + CInt(mySplit(3))
        iS4 = iS4 + CInt(mySplit(4))
        iTotal = CInt(mySplit(1)) + CInt(mySplit(2)) + CInt(mySplit(3)) + CInt(mySplit(4))
        
        '   Replace depending on status
        Select Case mySplit(0)
            Case "New"
                strNewText = Replace(strNewText, "NewSev0", mySplit(1))
                strNewText = Replace(strNewText, "NewSev1", mySplit(2))
                strNewText = Replace(strNewText, "NewSev2", mySplit(3))
                strNewText = Replace(strNewText, "NewSev3", mySplit(4))
                strNewText = Replace(strNewText, "NewSev4", iTotal)
            Case "Assigned"
                strNewText = Replace(strNewText, "AssSev0", mySplit(1))
                strNewText = Replace(strNewText, "AssSev1", mySplit(2))
                strNewText = Replace(strNewText, "AssSev2", mySplit(3))
                strNewText = Replace(strNewText, "AssSev3", mySplit(4))
                strNewText = Replace(strNewText, "AssSev4", iTotal)
            Case "Open"
                strNewText = Replace(strNewText, "OpenSev0", mySplit(1))
                strNewText = Replace(strNewText, "OpenSev1", mySplit(2))
                strNewText = Replace(strNewText, "OpenSev2", mySplit(3))
                strNewText = Replace(strNewText, "OpenSev3", mySplit(4))
                strNewText = Replace(strNewText, "OpenSev4", iTotal)
            Case "Reopen"
                strNewText = Replace(strNewText, "ReopenSev0", mySplit(1))
                strNewText = Replace(strNewText, "ReopenSev1", mySplit(2))
                strNewText = Replace(strNewText, "ReopenSev2", mySplit(3))
                strNewText = Replace(strNewText, "ReopenSev3", mySplit(4))
                strNewText = Replace(strNewText, "ReopenSev4", iTotal)
            Case "Failed Testing"
                strNewText = Replace(strNewText, "FailSev0", mySplit(1))
                strNewText = Replace(strNewText, "FailSev1", mySplit(2))
                strNewText = Replace(strNewText, "FailSev2", mySplit(3))
                strNewText = Replace(strNewText, "FailSev3", mySplit(4))
                strNewText = Replace(strNewText, "FailSev4", iTotal)
        End Select
        
        iTotal = 0
    Next
    '  Now write the total line
    iTotal = iS1 + iS2 + iS3 + iS4
    strNewText = Replace(strNewText, "iSO1", iS1)
    strNewText = Replace(strNewText, "iSO2", iS2)
    strNewText = Replace(strNewText, "iSO3", iS3)
    strNewText = Replace(strNewText, "iSO4", iS4)
    strNewText = Replace(strNewText, "iSOTotal", iTotal)
    
    '  Now the remainder
    For iSev = 5 To UBound(mySeverity)
       
        '  Split the record
        mySplit = Split(mySeverity(iSev), "|")
        ' Keep running total of each record
        iS1 = iS1 + CInt(mySplit(1))
        iS2 = iS2 + CInt(mySplit(2))
        iS3 = iS3 + CInt(mySplit(3))
        iS4 = iS4 + CInt(mySplit(4))
        iTotal = CInt(mySplit(1)) + CInt(mySplit(2)) + CInt(mySplit(3)) + CInt(mySplit(4))
        Select Case mySplit(0)
            Case "Fixed"
                strNewText = Replace(strNewText, "FixedSev0", mySplit(1))
                strNewText = Replace(strNewText, "FixedSev1", mySplit(2))
                strNewText = Replace(strNewText, "FixedSev2", mySplit(3))
                strNewText = Replace(strNewText, "FixedSev3", mySplit(4))
                strNewText = Replace(strNewText, "FixedSev4", iTotal)
            Case "Ready For Testing"
                strNewText = Replace(strNewText, "ReadySev0", mySplit(1))
                strNewText = Replace(strNewText, "ReadySev1", mySplit(2))
                strNewText = Replace(strNewText, "ReadySev2", mySplit(3))
                strNewText = Replace(strNewText, "ReadySev3", mySplit(4))
                strNewText = Replace(strNewText, "ReadySev4", iTotal)
            Case "Tested"
                strNewText = Replace(strNewText, "TestedSev0", mySplit(1))
                strNewText = Replace(strNewText, "TestedSev1", mySplit(2))
                strNewText = Replace(strNewText, "TestedSev2", mySplit(3))
                strNewText = Replace(strNewText, "TestedSev3", mySplit(4))
                strNewText = Replace(strNewText, "TestedSev4", iTotal)
            Case "Duplicate"
                strNewText = Replace(strNewText, "DupSev0", mySplit(1))
                strNewText = Replace(strNewText, "DupSev1", mySplit(2))
                strNewText = Replace(strNewText, "DupSev2", mySplit(3))
                strNewText = Replace(strNewText, "DupSev3", mySplit(4))
                strNewText = Replace(strNewText, "DupSev4", iTotal)
            Case "Rejected"
                strNewText = Replace(strNewText, "RejSev0", mySplit(1))
                strNewText = Replace(strNewText, "RejSev1", mySplit(2))
                strNewText = Replace(strNewText, "RejSev2", mySplit(3))
                strNewText = Replace(strNewText, "RejSev3", mySplit(4))
                strNewText = Replace(strNewText, "RejSev4", iTotal)
            Case "On Hold"
                strNewText = Replace(strNewText, "HoldSev0", mySplit(1))
                strNewText = Replace(strNewText, "HoldSev1", mySplit(2))
                strNewText = Replace(strNewText, "HoldSev2", mySplit(3))
                strNewText = Replace(strNewText, "HoldSev3", mySplit(4))
                strNewText = Replace(strNewText, "HoldSev4", iTotal)
            Case "Closed"
                strNewText = Replace(strNewText, "ClosedSev0", mySplit(1))
                strNewText = Replace(strNewText, "ClosedSev1", mySplit(2))
                strNewText = Replace(strNewText, "ClosedSev2", mySplit(3))
                strNewText = Replace(strNewText, "ClosedSev3", mySplit(4))
                strNewText = Replace(strNewText, "ClosedSev4", iTotal)
        End Select
        iTotal = 0
    Next
    '  Now write the total line
    iTotal = iS1 + iS2 + iS3 + iS4
    strNewText = Replace(strNewText, "iS1", iS1)
    strNewText = Replace(strNewText, "iS2", iS2)
    strNewText = Replace(strNewText, "iS3", iS3)
    strNewText = Replace(strNewText, "iS4", iS4)
    strNewText = Replace(strNewText, "iSTotal", iTotal)
    
    '  Now priority
    iS1 = 0
    iS2 = 0
    iS3 = 0
    iS4 = 0
    '  Write out the first four status
    For iSev = 0 To 4
        '  Split the record
        mySplit = Split(myPriority(iSev), "|")
        '  Keep running total of each status
        iS1 = iS1 + CInt(mySplit(1))
        iS2 = iS2 + CInt(mySplit(2))
        iS3 = iS3 + CInt(mySplit(3))
        iS4 = iS4 + CInt(mySplit(4))
        iTotal = CInt(mySplit(1)) + CInt(mySplit(2)) + CInt(mySplit(3)) + CInt(mySplit(4))
        
        '   Replace depending on status
        Select Case mySplit(0)
            Case "New"
                strNewText = Replace(strNewText, "NewPr0", mySplit(1))
                strNewText = Replace(strNewText, "NewPr1", mySplit(2))
                strNewText = Replace(strNewText, "NewPr2", mySplit(3))
                strNewText = Replace(strNewText, "NewPr3", mySplit(4))
                strNewText = Replace(strNewText, "NewPr4", iTotal)
            Case "Assigned"
                strNewText = Replace(strNewText, "AssPr0", mySplit(1))
                strNewText = Replace(strNewText, "AssPr1", mySplit(2))
                strNewText = Replace(strNewText, "AssPr2", mySplit(3))
                strNewText = Replace(strNewText, "AssPr3", mySplit(4))
                strNewText = Replace(strNewText, "AssPr4", iTotal)
            Case "Open"
                strNewText = Replace(strNewText, "OpenPr0", mySplit(1))
                strNewText = Replace(strNewText, "OpenPr1", mySplit(2))
                strNewText = Replace(strNewText, "OpenPr2", mySplit(3))
                strNewText = Replace(strNewText, "OpenPr3", mySplit(4))
                strNewText = Replace(strNewText, "OpenPr4", iTotal)
            Case "Reopen"
                strNewText = Replace(strNewText, "ReopenPr0", mySplit(1))
                strNewText = Replace(strNewText, "ReopenPr1", mySplit(2))
                strNewText = Replace(strNewText, "ReopenPr2", mySplit(3))
                strNewText = Replace(strNewText, "ReopenPr3", mySplit(4))
                strNewText = Replace(strNewText, "ReopenPr4", iTotal)
             Case "Failed Testing"
                strNewText = Replace(strNewText, "FailPr0", mySplit(1))
                strNewText = Replace(strNewText, "FailPr1", mySplit(2))
                strNewText = Replace(strNewText, "FailPr2", mySplit(3))
                strNewText = Replace(strNewText, "FailPr3", mySplit(4))
                strNewText = Replace(strNewText, "FailPr4", iTotal)
        End Select
        
        iTotal = 0
    Next
    '  Now write the total line
    iTotal = iS1 + iS2 + iS3 + iS4
    '  Now the remainder
    strNewText = Replace(strNewText, "iPO1", iS1)
    strNewText = Replace(strNewText, "iPO2", iS2)
    strNewText = Replace(strNewText, "iPO3", iS3)
    strNewText = Replace(strNewText, "iPO4", iS4)
    strNewText = Replace(strNewText, "iPOTotal", iTotal)
    
    For iSev = 5 To UBound(myPriority)

        '  Split the record
        mySplit = Split(myPriority(iSev), "|")
        ' Keep running total of each record
        iS1 = iS1 + CInt(mySplit(1))
        iS2 = iS2 + CInt(mySplit(2))
        iS3 = iS3 + CInt(mySplit(3))
        iS4 = iS4 + CInt(mySplit(4))
        iTotal = CInt(mySplit(1)) + CInt(mySplit(2)) + CInt(mySplit(3)) + CInt(mySplit(4))
        Select Case mySplit(0)
            Case "Fixed"
                strNewText = Replace(strNewText, "FixedPr0", mySplit(1))
                strNewText = Replace(strNewText, "FixedPr1", mySplit(2))
                strNewText = Replace(strNewText, "FixedPr2", mySplit(3))
                strNewText = Replace(strNewText, "FixedPr3", mySplit(4))
                strNewText = Replace(strNewText, "FixedPr4", iTotal)
            Case "Ready For Testing"
                strNewText = Replace(strNewText, "ReadyPr0", mySplit(1))
                strNewText = Replace(strNewText, "ReadyPr1", mySplit(2))
                strNewText = Replace(strNewText, "ReadyPr2", mySplit(3))
                strNewText = Replace(strNewText, "ReadyPr3", mySplit(4))
                strNewText = Replace(strNewText, "ReadyPr4", iTotal)
            Case "Tested"
                strNewText = Replace(strNewText, "TestedPr0", mySplit(1))
                strNewText = Replace(strNewText, "TestedPr1", mySplit(2))
                strNewText = Replace(strNewText, "TestedPr2", mySplit(3))
                strNewText = Replace(strNewText, "TestedPr3", mySplit(4))
                strNewText = Replace(strNewText, "TestedPr4", iTotal)
            Case "Duplicate"
                strNewText = Replace(strNewText, "DupPr0", mySplit(1))
                strNewText = Replace(strNewText, "DupPr1", mySplit(2))
                strNewText = Replace(strNewText, "DupPr2", mySplit(3))
                strNewText = Replace(strNewText, "DupPr3", mySplit(4))
                strNewText = Replace(strNewText, "DupPr4", iTotal)
            Case "Rejected"
                strNewText = Replace(strNewText, "RejPr0", mySplit(1))
                strNewText = Replace(strNewText, "RejPr1", mySplit(2))
                strNewText = Replace(strNewText, "RejPr2", mySplit(3))
                strNewText = Replace(strNewText, "RejPr3", mySplit(4))
                strNewText = Replace(strNewText, "RejPr4", iTotal)
            Case "On Hold"
                strNewText = Replace(strNewText, "HoldPr0", mySplit(1))
                strNewText = Replace(strNewText, "HoldPr1", mySplit(2))
                strNewText = Replace(strNewText, "HoldPr2", mySplit(3))
                strNewText = Replace(strNewText, "HoldPr3", mySplit(4))
                strNewText = Replace(strNewText, "HoldPr4", iTotal)
            Case "Closed"
                strNewText = Replace(strNewText, "ClosedPr0", mySplit(1))
                strNewText = Replace(strNewText, "ClosedPr1", mySplit(2))
                strNewText = Replace(strNewText, "ClosedPr2", mySplit(3))
                strNewText = Replace(strNewText, "ClosedPr3", mySplit(4))
                strNewText = Replace(strNewText, "ClosedPr4", iTotal)
        End Select
        iTotal = 0
    Next
    '  Now write the total line
    iTotal = iS1 + iS2 + iS3 + iS4
    strNewText = Replace(strNewText, "iP1", iS1)
    strNewText = Replace(strNewText, "iP2", iS2)
    strNewText = Replace(strNewText, "iP3", iS3)
    strNewText = Replace(strNewText, "iP4", iS4)
    strNewText = Replace(strNewText, "iPTotal", iTotal)
    
    '   Write the new text to the asp file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForWriting, True)
    myFile.WriteLine strNewText
    myFile.Close
    
End Function
Public Function DefectRootCauseBySeverity()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    Dim arrRootCause(3)
    Dim strRootCause As String
    Dim strSeverity As String
    Dim iCode As Integer, iConfig As Integer, iData As Integer, iEnviron As Integer, iReq As Integer, iTestCase As Integer
    Dim iCodeS1 As Integer, iConfigS1 As Integer, iDataS1 As Integer, iEnvironS1 As Integer, iReqS1 As Integer, iTestCaseS1 As Integer, iDuplicateS1 As Integer, iRejectedS1 As Integer
    Dim iCodeS2 As Integer, iConfigS2 As Integer, iDataS2 As Integer, iEnvironS2 As Integer, iReqS2 As Integer, iTestCaseS2 As Integer, iDuplicateS2 As Integer, iRejectedS2 As Integer
    Dim iCodeS3 As Integer, iConfigS3 As Integer, iDataS3 As Integer, iEnvironS3 As Integer, iReqS3 As Integer, iTestCaseS3 As Integer, iDuplicateS3 As Integer, iRejectedS3 As Integer
    Dim iCodeS4 As Integer, iConfigS4 As Integer, iDataS4 As Integer, iEnvironS4 As Integer, iReqS4 As Integer, iTestCaseS4 As Integer, iDuplicateS4 As Integer, iRejectedS4 As Integer
    
    '   Create the filter
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    
    '   If the root cause is empty then ignore from the calculations.
    tdcBugFilter.Filter(strRootCauseLabel) = "Not " & Chr(34) & Chr(34)
    
    '   Get the list
    Set colBugList = tdcBugFilter.NewList
    
    For Each objBug In colBugList
        If objBug.Status = "Fixed" Or objBug.Status = "Closed" Or objBug.Status = "Tested" Then
            strRootCause = objBug.Field(strRootCauseLabel)
            strSeverity = objBug.Field("BG_SEVERITY")
            Select Case strRootCause
                Case "Code"
                    Select Case strSeverity
                        Case "1-Critical"
                            iCodeS1 = iCodeS1 + 1
                        Case "2-High"
                            iCodeS2 = iCodeS2 + 1
                        Case "3-Medium"
                            iCodeS3 = iCodeS3 + 1
                        Case "4-Low"
                            iCodeS4 = iCodeS4 + 1
                        Case Else
                            iCode = iCode + 1
                    End Select
                Case "Configuration"
                    Select Case strSeverity
                        Case "1-Critical"
                            iConfigS1 = iConfigS1 + 1
                        Case "2-High"
                            iConfigS2 = iConfigS2 + 1
                        Case "3-Medium"
                            iConfigS3 = iConfigS3 + 1
                        Case "4-Low"
                            iConfigS4 = iConfigS4 + 1
                        Case Else
                            iConfig = iConfig + 1
                    End Select
                Case "Data"
                    Select Case strSeverity
                        Case "1-Critical"
                            iDataS1 = iDataS1 + 1
                        Case "2-High"
                            iDataS2 = iDataS2 + 1
                        Case "3-Medium"
                            iDataS3 = iDataS3 + 1
                        Case "4-Low"
                            iDataS4 = iDataS4 + 1
                        Case Else
                            iData = iData + 1
                    End Select
                Case "Environment"
                    Select Case strSeverity
                        Case "1-Critical"
                            iEnvironS1 = iEnvironS1 + 1
                        Case "2-High"
                            iEnvironS2 = iEnvironS2 + 1
                        Case "3-Medium"
                            iEnvironS3 = iEnvironS3 + 1
                        Case "4-Low"
                            iEnvironS4 = iEnvironS4 + 1
                        Case Else
                            iEnviron = iEnviron + 1
                    End Select
                Case "Requirement"
                    Select Case strSeverity
                        Case "1-Critical"
                            iReqS1 = iReqS1 + 1
                        Case "2-High"
                            iReqS2 = iReqS2 + 1
                        Case "3-Medium"
                            iReqS3 = iReqS3 + 1
                        Case "4-Low"
                            iReqS4 = iReqS4 + 1
                        Case Else
                            iReq = iReq + 1
                    End Select
                Case "Test Case"
                    Select Case strSeverity
                        Case "1-Critical"
                            iTestCaseS1 = iTestCaseS1 + 1
                        Case "2-High"
                            iTestCaseS2 = iTestCaseS2 + 1
                        Case "3-Medium"
                            iTestCaseS3 = iTestCaseS3 + 1
                        Case "4-Low"
                            iTestCaseS4 = iTestCaseS4 + 1
                        Case Else
                            iTestCase = iTestCase + 1
                    End Select
                Case "Rejected"
                    Select Case strSeverity
                        Case "1-Critical"
                            iRejectedS1 = iRejectedS1 + 1
                        Case "2-High"
                            iRejectedS2 = iRejectedS2 + 1
                        Case "3-Medium"
                            iRejectedS3 = iRejectedS3 + 1
                        Case "4-Low"
                            iRejectedS4 = iRejectedS4 + 1
                        Case Else
                            iRejected = iTestCase + 1
                    End Select
                Case "Duplicate"
                    Select Case strSeverity
                        Case "1-Critical"
                            iDuplicateS1 = iDuplicateS1 + 1
                        Case "2-High"
                            iDuplicateS2 = iDuplicateS2 + 1
                        Case "3-Medium"
                            iDuplicateS3 = iDuplicateS3 + 1
                        Case "4-Low"
                            iDuplicateS4 = iDuplicateS4 + 1
                        Case Else
                            iDuplicate = iDuplicate + 1
                    End Select
            End Select
        End If
    Next
    
    '   Put them all into an array
    arrRootCause(0) = iCodeS1 & "|" & iConfigS1 & "|" & iDataS1 & "|" & iEnvironS1 & "|" & iReqS1 & "|" & iTestCaseS1 & "|" & iRejectedS1 & "|" & iDuplicateS1
    arrRootCause(1) = iCodeS2 & "|" & iConfigS2 & "|" & iDataS2 & "|" & iEnvironS2 & "|" & iReqS2 & "|" & iTestCaseS2 & "|" & iRejectedS2 & "|" & iDuplicateS2
    arrRootCause(2) = iCodeS3 & "|" & iConfigS3 & "|" & iDataS3 & "|" & iEnvironS3 & "|" & iReqS3 & "|" & iTestCaseS3 & "|" & iRejectedS3 & "|" & iDuplicateS3
    arrRootCause(3) = iCodeS4 & "|" & iConfigS4 & "|" & iDataS4 & "|" & iEnvironS4 & "|" & iReqS4 & "|" & iTestCaseS4 & "|" & iRejectedS4 & "|" & iDuplicateS4
    
    DefectRootCauseBySeverity = arrRootCause
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function GetDefectRootCause()
    
    ' Count the critical defects by root cause and add them to the data sheet.
    arrRootCause = DefectRootCauseBySeverity()

    '   Build the chart stuff into the template
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "DefectRootCauseTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectRootCause.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "1-CriticalPoints") > 0 Then
            '   Split the file
            mySplit = Split(arrRootCause(0), "|")
            '   Write the values
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Code" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "Configuration" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "Data" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "Environment" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "Requirements" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "Test Case" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "Rejected" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "Duplicate" & Chr(34) & " />"
        Else
            If InStr(1, rc, "2-HighPoints") > 0 Then
                '   Split the file
                mySplit = Split(arrRootCause(1), "|")
                '   Write the values
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Code" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "Configuration" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "Data" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "Environment" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "Requirements" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "Test Case" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "Rejected" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "Duplicate" & Chr(34) & " />"
            Else
                If InStr(1, rc, "3-MediumPoints") > 0 Then
                    '   Split the file
                    mySplit = Split(arrRootCause(2), "|")
                    '   Write the values
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Code" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "Configuration" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "Data" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "Environment" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "Requirements" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "Test Case" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "Rejected" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "Duplicate" & Chr(34) & " />"
                Else
                    If InStr(1, rc, "4-LowPoints") > 0 Then
                            '   Split the file
                            mySplit = Split(arrRootCause(3), "|")
                            '   Write the values
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Code" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "Configuration" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "Data" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "Environment" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "Requirements" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "Test Case" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "Rejected" & Chr(34) & " />"
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "Duplicate" & Chr(34) & " />"
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
    
    '   See if we need to amend or remove the y axis interval
    iCount = 0
    mySplit = Split(arrRootCause(0))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                If CInt(NextEle) > iCount Then
                    iCount = CInt(NextEle)
                End If
            End If
        Next
    Next
    mySplit = Split(arrRootCause(1))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If CInt(NextEle) > iCount Then
                iCount = CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrRootCause(2))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If CInt(NextEle) > iCount Then
                iCount = CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrRootCause(3))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If CInt(NextEle) > iCount Then
                iCount = CInt(NextEle)
            End If
        Next
    Next
    
    '   Open the time open file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectRootCause.aspx", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    '   Look at the count to decide how we change the y axis interval
    Select Case iCount
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
            strText = Replace(strText, "ReplaceYAxis", "Interval=" & Chr(34) & "1" & Chr(34))
        Case Else
            strText = Replace(strText, "ReplaceYAxis", "")
    End Select
    '   Replace in the file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectRootCause.aspx", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Now open the table template for new - fixed
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectRootCauseTableTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    '   Replace all the values with our data
    
    '   Critical 1st
    mySplit = Split(arrRootCause(0), "|")
    strText = Replace(strText, "CodeCrit", mySplit(0))
    strText = Replace(strText, "ConfigCrit", mySplit(1))
    strText = Replace(strText, "DataCrit", mySplit(2))
    strText = Replace(strText, "EnvCrit", mySplit(3))
    strText = Replace(strText, "ReqCrit", mySplit(4))
    strText = Replace(strText, "TCCrit", mySplit(5))
    strText = Replace(strText, "RejCrit", mySplit(6))
    strText = Replace(strText, "DupCrit", mySplit(7))
    mySplit = Split(arrRootCause(1), "|")
    strText = Replace(strText, "CodeHigh", mySplit(0))
    strText = Replace(strText, "ConfigHigh", mySplit(1))
    strText = Replace(strText, "DataHigh", mySplit(2))
    strText = Replace(strText, "EnvHigh", mySplit(3))
    strText = Replace(strText, "ReqHigh", mySplit(4))
    strText = Replace(strText, "TCHigh", mySplit(5))
    strText = Replace(strText, "RejHigh", mySplit(6))
    strText = Replace(strText, "DupHigh", mySplit(7))
    mySplit = Split(arrRootCause(2), "|")
    strText = Replace(strText, "CodeMed", mySplit(0))
    strText = Replace(strText, "ConfigMed", mySplit(1))
    strText = Replace(strText, "DataMed", mySplit(2))
    strText = Replace(strText, "EnvMed", mySplit(3))
    strText = Replace(strText, "ReqMed", mySplit(4))
    strText = Replace(strText, "TCMed", mySplit(5))
    strText = Replace(strText, "RejMed", mySplit(6))
    strText = Replace(strText, "DupMed", mySplit(7))
    mySplit = Split(arrRootCause(3), "|")
    strText = Replace(strText, "CodeLow", mySplit(0))
    strText = Replace(strText, "ConfigLow", mySplit(1))
    strText = Replace(strText, "DataLow", mySplit(2))
    strText = Replace(strText, "EnvLow", mySplit(3))
    strText = Replace(strText, "ReqLow", mySplit(4))
    strText = Replace(strText, "TCLow", mySplit(5))
    strText = Replace(strText, "RejLow", mySplit(6))
    strText = Replace(strText, "DupLow", mySplit(7))
    
    '   Now open the file for new - fixed
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectRootCauseTable.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
End Function
Public Function GetDefectsFromOtherPhases()
Dim a As Integer
a = 0

    ' Count the critical defects by test phase and add them to the data sheet.
    arrTestPhaseKeys = DefectTestPhaseBySeverity()
    QSortInPlace arrTestPhaseKeys
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "DefectTestPhaseTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectTestPhase.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "1-CriticalPoints") > 0 Then
            '   Loop round our array
            For i = 0 To UBound(arrTestPhaseKeys)
                '   Split the file
                mySplit = Split(arrTestPhaseKeys(i), "|")
                If mySplit(0) = "1-Critical" Then
                    '   Write the values
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestPhaseBySeverityDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                End If
            Next
        Else
            If InStr(1, rc, "2-HighPoints") > 0 Then
                '   Loop round our array
                For i = 0 To UBound(arrTestPhaseKeys)
                    '   Split the file
                    mySplit = Split(arrTestPhaseKeys(i), "|")
                    If mySplit(0) = "2-High" Then
                        '   Write the values
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestPhaseBySeverityDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                    End If
                Next
            Else
                If InStr(1, rc, "3-MediumPoints") > 0 Then
                    '   Loop round our array
                    For i = 0 To UBound(arrTestPhaseKeys)
                        '   Split the file
                        mySplit = Split(arrTestPhaseKeys(i), "|")
                        If mySplit(0) = "3-Medium" Then
                            '   Write the values
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestPhaseBySeverityDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                        End If
                    Next
                Else
                    If InStr(1, rc, "4-LowPoints") > 0 Then
                        '   Loop round our array
                        For i = 0 To UBound(arrTestPhaseKeys)
                            '   Split the file
                            mySplit = Split(arrTestPhaseKeys(i), "|")
                            If mySplit(0) = "4-Low" Then
                                '   Write the values
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objTestPhaseBySeverityDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
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
    
End Function
Public Function DefectTestPhaseBySeverity()

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    Dim strSeverity As String
    Dim lstTestPhaseList As list
    Dim intTestPhaseCount As Integer
    Dim iCount As Integer
    Dim arrSeverity(3)
    
    arrSeverity(0) = "1-Critical"
    arrSeverity(1) = "2-High"
    arrSeverity(2) = "3-Medium"
    arrSeverity(3) = "4-Low"
    
    ' Create the dictionary object
    Set objTestPhaseBySeverityDictionary = New Dictionary
    
    '   Get the list of test phases within this project
    Set lstTestPhaseList = GetListValues("Test Phase")
    
    intTestPhaseCount = 1
    Do Until intTestPhaseCount = lstTestPhaseList.Count + 1
        objTestPhaseBySeverityDictionary.Add arrSeverity(0) & "|" & lstTestPhaseList.Item(intTestPhaseCount), 0
        objTestPhaseBySeverityDictionary.Add arrSeverity(1) & "|" & lstTestPhaseList.Item(intTestPhaseCount), 0
        objTestPhaseBySeverityDictionary.Add arrSeverity(2) & "|" & lstTestPhaseList.Item(intTestPhaseCount), 0
        objTestPhaseBySeverityDictionary.Add arrSeverity(3) & "|" & lstTestPhaseList.Item(intTestPhaseCount), 0
        intTestPhaseCount = intTestPhaseCount + 1
    Loop
    
    '   Set up the filter
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter("BG_STATUS") = " Not (" & Chr(34) & "Closed" & Chr(34) & " Or " & Chr(34) & "Rejected" & Chr(34) & " Or " & Chr(34) & "Duplicate" & Chr(34) & " Or " & Chr(34) & "On Hold" & Chr(34) & " Or " & Chr(34) & "Tested" & Chr(34) & ")"
    tdcBugFilter.Filter("BG_DETECTION_DATE") = "<= " & TodaysDate
    
    '   Create the list
    Set colBugList = tdcBugFilter.NewList
    
    '   Go through each test phase in the list and see if we have any stats for it
    For Each objBug In colBugList
        strThisTestPhase = objBug.Field(strTestPhaseBugLabel)
        strSeverity = objBug.Field("BG_SEVERITY")
        intCount = CInt(objTestPhaseBySeverityDictionary.Item(strSeverity & "|" & strThisTestPhase))
        objTestPhaseBySeverityDictionary.Item(strSeverity & "|" & strThisTestPhase) = intCount + 1
    Next
    
    DefectTestPhaseBySeverity = objTestPhaseBySeverityDictionary.Keys
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function GetDailyDefects() As Integer
Dim intDefectsAccepted As Integer
Dim intDefectsRaised As Integer
Dim intDefectsClosed As Integer
Dim intDefectsFixed As Integer
Dim intColumn As Integer
Dim strProjectLabel As String
Dim strDateFilter As String
Dim strTestedOnDate As String
Dim strFixedOnDate As String
Dim strSDate As Date
Dim arrWeekends
Dim iTot As Integer
Dim iFix As Integer
Dim iTest As Integer
Dim iDum As Integer
Dim blnErrorPage As Boolean
iTot = 0
iAcc = 0
iFix = 0
iTest = 0
iDum = 0
iClosed = 0
ReDim Preserve arrDummy(iDum)
arrDummy(iDum) = "Project Start|0"
ReDim Preserve arrTotal(iTot)
arrTotal(iTot) = "Project Start|0"
ReDim Preserve arrAccepted(iAcc)
arrAccepted(iAcc) = "Project Start|0"
ReDim Preserve arrFixed(iFix)
arrFixed(iFix) = "Project Start|0"
ReDim Preserve arrTested(iTest)
arrTested(iTest) = "Project Start|0"
ReDim Preserve arrClosed(iClosed)
arrClosed(iClosed) = "Project Start|0"

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    intCurrentRow = 1
    
    intColumn = 2
    
    '   Find the 1st bug date
    strSDate = FindFirstBugDate
    
    '   See if we've got no bugs
    If strSDate = "00:00:00" Then
        dateLastWeek = DateAdd("d", -7, TodaysDate)
        blnErrorPage = True
        GoTo WriteHTML
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
    arrMonths = FindMonths(strSDate, datePreviousMonth)
    If arrMonths(0) <> "No Month" Then
        arrWeekends = FindWeekends(arrMonths(UBound(arrMonths)), dateLastWeek)
    Else
        If datePreviousMonth > strSDate Then
            arrWeekends = FindWeekends(datePreviousMonth, dateLastWeek)
        Else
            arrWeekends = FindWeekends(strSDate, dateLastWeek)
        End If
    End If

    '   Firstly get the values for the very 1st date
    strDateFilter = "= " & strSDate
    '   Get the values
    intDefectsRaised = CountDefectsByDate("New", strDateFilter)
    intDefectsAccepted = CountDefectsByDate("Accepted", strDateFilter)
    intDefectsFixed = CountDefectsByDate("Fixed", strDateFilter)
    intDefectsTested = CountDefectsByDate("Tested", strDateFilter)
    intDefectsClosed = CountDefectsByDate("Closed", strDateFilter)
    '   Add the date and values to the arrays
    iTot = iTot + 1
    ReDim Preserve arrTotal(iTot)
    arrTotal(iTot) = strSDate & "|" & intDefectsRaised
    iAcc = iAcc + 1
    ReDim Preserve arrAccepted(iAcc)
    arrAccepted(iAcc) = strSDate & "|" & intDefectsAccepted
    iFix = iFix + 1
    ReDim Preserve arrFixed(iFix)
    arrFixed(iFix) = strSDate & "|" & intDefectsFixed
    iTest = iTest + 1
    ReDim Preserve arrTested(iTest)
    arrTested(iTest) = strSDate & "|" & intDefectsTested
    iClosed = iClosed + 1
    ReDim Preserve arrClosed(iClosed)
    arrClosed(iClosed) = strSDate & "|" & intDefectsClosed
    iDum = iDum + 1
    ReDim Preserve arrDummy(iDum)
    arrDummy(iDum) = strSDate & "|0"
    '   Move the day on
    strSDate = strSDate + 1

    '   Loop round the Months getting the data for each month prior
    If arrMonths(0) <> "No Month" Then
        For Each Ele In arrMonths
            intDefectsRaised = 0
            intDefectsAccepted = 0
            intDefectsFixed = 0
            intDefectsTested = 0
            intDefectsClosed = 0
            
            '   Set the date filter
            strDateFilter = ">= " & strSDate & " And < " & Ele
            
            '   Get the values
            intDefectsRaised = CountDefectsByDate("New", strDateFilter)
            intDefectsAccepted = CountDefectsByDate("Accepted", strDateFilter)
            intDefectsFixed = CountDefectsByDate("Fixed", strDateFilter)
            intDefectsTested = CountDefectsByDate("Tested", strDateFilter)
            intDefectsClosed = CountDefectsByDate("Closed", strDateFilter)
            
            '   Add the date and values to the arrays
            iTot = iTot + 1
            ReDim Preserve arrTotal(iTot)
            arrTotal(iTot) = strSDate & "|" & intDefectsRaised
            iAcc = iAcc + 1
            ReDim Preserve arrAccepted(iAcc)
            arrAccepted(iAcc) = strSDate & "|" & intDefectsAccepted
            iFix = iFix + 1
            ReDim Preserve arrFixed(iFix)
            arrFixed(iFix) = strSDate & "|" & intDefectsFixed
            iTest = iTest + 1
            ReDim Preserve arrTested(iTest)
            arrTested(iTest) = strSDate & "|" & intDefectsTested
            iClosed = iClosed + 1
            ReDim Preserve arrClosed(iClosed)
            arrClosed(iClosed) = strSDate & "|" & intDefectsClosed
            iDum = iDum + 1
            ReDim Preserve arrDummy(iDum)
            arrDummy(iDum) = strSDate & "|0"
            
            '   Make these cumulative
            If iTot > 0 Then
                myoldsplit = Split(arrTotal(iTot - 1), "|")
                mycurrsplit = Split(arrTotal(iTot), "|")
                iNewTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrTotal(iTot) = mycurrsplit(0) & "|" & iNewTot
            End If
            If iAcc > 0 Then
                myoldsplit = Split(arrAccepted(iAcc - 1), "|")
                mycurrsplit = Split(arrAccepted(iAcc), "|")
                iNewAcc = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrAccepted(iAcc) = mycurrsplit(0) & "|" & iNewAcc
            End If
            If iFix > 0 Then
                myoldsplit = Split(arrFixed(iFix - 1), "|")
                mycurrsplit = Split(arrFixed(iFix), "|")
                iNewFix = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrFixed(iFix) = mycurrsplit(0) & "|" & iNewFix
            End If
            If iTest > 0 Then
                myoldsplit = Split(arrTested(iTest - 1), "|")
                mycurrsplit = Split(arrTested(iTest), "|")
                iNewTest = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrTested(iTest) = mycurrsplit(0) & "|" & iNewTest
            End If
            If iClosed > 0 Then
                myoldsplit = Split(arrClosed(iClosed - 1), "|")
                mycurrsplit = Split(arrClosed(iClosed), "|")
                iNewClosed = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrClosed(iClosed) = mycurrsplit(0) & "|" & iNewClosed
            End If
            
            ' Set the start date to this element
            strSDate = Ele
           
        Next
    End If
    If arrWeekends(0) <> "No Weeks" Then
        For Each Ele In arrWeekends
            intDefectsRaised = 0
            intDefectsAccepted = 0
            intDefectsFixed = 0
            intDefectsTested = 0
            intDefectsClosed = 0
            
            '   Set the date filter
            strDateFilter = ">= " & strSDate & " And < " & Ele
            
            '   Get the values
            intDefectsRaised = CountDefectsByDate("New", strDateFilter)
            intDefectsAccepted = CountDefectsByDate("Accepted", strDateFilter)
            intDefectsFixed = CountDefectsByDate("Fixed", strDateFilter)
            intDefectsTested = CountDefectsByDate("Tested", strDateFilter)
            intDefectsClosed = CountDefectsByDate("Closed", strDateFilter)
            
            '   Add the date and values to the arrays
            iTot = iTot + 1
            ReDim Preserve arrTotal(iTot)
            arrTotal(iTot) = strSDate & "|" & intDefectsRaised
            iAcc = iAcc + 1
            ReDim Preserve arrAccepted(iAcc)
            arrAccepted(iAcc) = strSDate & "|" & intDefectsAccepted
            iFix = iFix + 1
            ReDim Preserve arrFixed(iFix)
            arrFixed(iFix) = strSDate & "|" & intDefectsFixed
            iTest = iTest + 1
            ReDim Preserve arrTested(iTest)
            arrTested(iTest) = strSDate & "|" & intDefectsTested
            iClosed = iClosed + 1
            ReDim Preserve arrClosed(iClosed)
            arrClosed(iClosed) = strSDate & "|" & intDefectsClosed
            iDum = iDum + 1
            ReDim Preserve arrDummy(iDum)
            arrDummy(iDum) = strSDate & "|0"
            
            '   Make these cumulative
            If iTot > 0 Then
                myoldsplit = Split(arrTotal(iTot - 1), "|")
                mycurrsplit = Split(arrTotal(iTot), "|")
                iNewTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrTotal(iTot) = mycurrsplit(0) & "|" & iNewTot
            End If
            If iAcc > 0 Then
                myoldsplit = Split(arrAccepted(iAcc - 1), "|")
                mycurrsplit = Split(arrAccepted(iAcc), "|")
                iNewAcc = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrAccepted(iAcc) = mycurrsplit(0) & "|" & iNewAcc
            End If
            If iFix > 0 Then
                myoldsplit = Split(arrFixed(iFix - 1), "|")
                mycurrsplit = Split(arrFixed(iFix), "|")
                iNewFix = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrFixed(iFix) = mycurrsplit(0) & "|" & iNewFix
            End If
            If iTest > 0 Then
                myoldsplit = Split(arrTested(iTest - 1), "|")
                mycurrsplit = Split(arrTested(iTest), "|")
                iNewTest = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrTested(iTest) = mycurrsplit(0) & "|" & iNewTest
            End If
            If iClosed > 0 Then
                myoldsplit = Split(arrClosed(iClosed - 1), "|")
                mycurrsplit = Split(arrClosed(iClosed), "|")
                iNewClosed = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
                arrClosed(iClosed) = mycurrsplit(0) & "|" & iNewClosed
            End If
            
            ' Set the start date to this element
            strSDate = Ele
        Next
    End If
    '   Now do the last lot of daily stats
    Do While strSDate <= TodaysDate
    
        '   Set the date filter
        strDateFilter = strSDate
        
        '   Get the values
        intDefectsRaised = 0
        intDefectsRaised = CountDefectsByDate("New", strDateFilter)
        
        intDefectsAccepted = 0
        intDefectsAccepted = CountDefectsByDate("Accepted", strDateFilter)
    
        intDefectsFixed = 0
        intDefectsFixed = CountDefectsByDate("Fixed", strDateFilter)
        
        intDefectsTested = 0
        intDefectsTested = CountDefectsByDate("Tested", strDateFilter)
        
        intDefectsClosed = 0
        intDefectsClosed = CountDefectsByDate("Closed", strDateFilter)
        
        '   Drop if a weekend with no data
        If (Weekday(strSDate) = 7 Or Weekday(strSDate) = 1) _
            And intDefectsRaised = 0 _
            And intDefectsAccepted = 0 _
            And intDefectsFixed = 0 _
            And intDefectsTested = 0 _
            And intDefectsClosed = 0 _
            Then
            GoTo ExitLoop
        End If
        
        '   Add the date and values to the arrays
        iTot = iTot + 1
        ReDim Preserve arrTotal(iTot)
        arrTotal(iTot) = strSDate & "|" & intDefectsRaised
        iAcc = iAcc + 1
        ReDim Preserve arrAccepted(iAcc)
        arrAccepted(iAcc) = strSDate & "|" & intDefectsAccepted
        iFix = iFix + 1
        ReDim Preserve arrFixed(iFix)
        arrFixed(iFix) = strSDate & "|" & intDefectsFixed
        iTest = iTest + 1
        ReDim Preserve arrTested(iTest)
        arrTested(iTest) = strSDate & "|" & intDefectsTested
        iClosed = iClosed + 1
        ReDim Preserve arrClosed(iClosed)
        arrClosed(iClosed) = strSDate & "|" & intDefectsClosed
        iDum = iDum + 1
        ReDim Preserve arrDummy(iDum)
        arrDummy(iDum) = strSDate & "|0"
            
        '   Make these cumulative
        If iTot > 0 Then
            myoldsplit = Split(arrTotal(iTot - 1), "|")
            mycurrsplit = Split(arrTotal(iTot), "|")
            iNewTot = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrTotal(iTot) = mycurrsplit(0) & "|" & iNewTot
        End If
        If iAcc > 0 Then
            myoldsplit = Split(arrAccepted(iAcc - 1), "|")
            mycurrsplit = Split(arrAccepted(iAcc), "|")
            iNewAcc = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrAccepted(iAcc) = mycurrsplit(0) & "|" & iNewAcc
        End If
        If iFix > 0 Then
            myoldsplit = Split(arrFixed(iFix - 1), "|")
            mycurrsplit = Split(arrFixed(iFix), "|")
            iNewFix = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrFixed(iFix) = mycurrsplit(0) & "|" & iNewFix
        End If
        If iTest > 0 Then
            myoldsplit = Split(arrTested(iTest - 1), "|")
            mycurrsplit = Split(arrTested(iTest), "|")
            iNewTest = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrTested(iTest) = mycurrsplit(0) & "|" & iNewTest
        End If
        If iClosed > 0 Then
            myoldsplit = Split(arrClosed(iClosed - 1), "|")
            mycurrsplit = Split(arrClosed(iClosed), "|")
            iNewClosed = CInt(myoldsplit(1)) + CInt(mycurrsplit(1))
            arrClosed(iClosed) = mycurrsplit(0) & "|" & iNewClosed
        End If
        
ExitLoop:
        '   Set the date to one date ahead
        strSDate = strSDate + 1
    Loop

    '   See if we've got any test sets and so an end date
    If blnNoTestSets = True Then
        dateEndDate = TodaysDate
    End If
    
    '   Now get just the planned values out into the future if required
    If dateEndDate > TodaysDate Then
        iDays = DateDiff("d", TodaysDate, dateEndDate)
        If iDays <= 30 Then
            strSDate = DateAdd("d", 1, TodaysDate)
            Do While strSDate <= dateEndDate
                iDum = iDum + 1
                ReDim Preserve arrDummy(iDum)
                arrDummy(iDum) = strSDate & "|0"
                strSDate = strSDate + 1
            Loop
            GoTo WriteHTML
        End If
        strSDate = DateAdd("d", 1, TodaysDate)
        '   Get the date next week
        dateNextWeek = DateAdd("d", 7, TodaysDate)
        '   See if it's past our end date
        If dateNextWeek >= dateEndDate Then
            '   Get the last value reported cos we're just going to repeat this
            Do While strSDate <= dateEndDate
                iDum = iDum + 1
                ReDim Preserve arrDummy(iDum)
                arrDummy(iDum) = strSDate & "|0"
                strSDate = strSDate + 1
            Loop
            GoTo WriteHTML
        End If
        dateNextMonth = DateAdd("m", 1, dateNextWeek)
        '   See if it's past our end date
        If dateNextMonth >= dateEndDate Then
            arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
            If arrWeekends(0) <> "No Weeks" Then
                If arrWeekends(0) < dateEndDate Then
                    Do While strSDate <= dateNextWeek
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                    For Each Ele In arrWeekends
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = Ele & "|0"
                        strSDate = Ele
                    Next
                    strSDate = strSDate + 1
                    Do While strSDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                Else
                    Do While strSDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                End If
            Else
                Do While strSDate <= dateEndDate
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = strSDate & "|0"
                    strSDate = strSDate + 1
                Loop
                GoTo WriteHTML
            End If
        End If
        '   See how many days we're dealing with
        iDays = DateDiff("d", dateNextMonth, dateEndDate)
            If iDays < 0 Then
                GoTo WriteHTML
            End If
            If iDays <= 30 Then
                arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
                '   Write out the first weeks worth of days
                Do While strSDate <= dateNextWeek
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = strSDate & "|0"
                    strSDate = strSDate + 1
                Loop
                If arrWeekends(0) <> "No Weeks" Then
                        '   Now write out the remaining weeks
                        For Each Ele In arrWeekends
                                iDum = iDum + 1
                                ReDim Preserve arrDummy(iDum)
                                arrDummy(iDum) = Ele & "|0"
                                strSDate = Ele
                        Next
                End If
                '   Write out last days if we've not hit the last date yet
                If strSDate < dateEndDate Then
                    strSDate = strSDate + 1
                    Do While strSDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve arrDummy(iDum)
                        arrDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                End If
            Else
                arrMonths = FindMonths(dateNextMonth, dateEndDate)
                arrWeekends = FindWeekends(dateNextWeek, dateNextMonth)
                Do While strSDate <= dateNextWeek
                    iDum = iDum + 1
                    ReDim Preserve arrDummy(iDum)
                    arrDummy(iDum) = strSDate & "|0"
                    strSDate = strSDate + 1
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
    
WriteHTML:

    '   See if we're defaulting to the error page or not
    If blnErrorPage = False Then
    
        '   Now write out the Detected vs Closed table info
        fso.CopyFile strTemplatePath & "NewDefectDailyTableTemplate.txt", strFolderPath & strPathandFileName & "-DefectDailyTable.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDailyTable.asp", ForAppending, True)
        
        '   Write out the table header details etc
        myFile.WriteLine "<table border=1 align=center>"
        myFile.WriteLine "<tr><th colspan=7 align=center>" & strHeader & "</th></tr>"
        myFile.WriteLine "<tr bgcolor='#CCCCCC'><td>&nbsp;</td><td>Accepted Defects</td><td>Fixed Defects</td><td>Tested Defects</td></tr>"
        iDetectedCount = UBound(arrAccepted)
        iDummyCount = UBound(arrDummy)
        If iDetectedCount = iDummyCount Then
            iTotal = iDummyCount
        Else
            iTotal = iDetectedCount
        End If
        i = 0
        Do
            '   Get the values from each of the arrays for this array element
            aSplit = Split(arrAccepted(i), "|")
            bSplit = Split(arrFixed(i), "|")
            cSplit = Split(arrTested(i), "|")
            dSplit = Split(arrDummy(i), "|")
            
            '   If we're on the first element then just default to project start and zeros
            If i = 0 Then
                myFile.WriteLine "<tr bgcolor ='#B5EAAA'><td>Project Start</td><td>0</td><td>0</td><td>0</td></tr>"
            Else
                If Weekday(aSplit(0)) <> 1 And Weekday(aSplit(0)) <> 7 Then
                    If IsEven(i) = True Then
                        myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
                    Else
                        myFile.WriteLine ("<tr bgcolor ='ivory'>")
                    End If
                    '   Write a row to the table
                    myFile.WriteLine "<td>" & aSplit(0) & "</td><td>" & aSplit(1) & "</td><td>" & bSplit(1) & "</td><td>" & cSplit(1) & "</td></tr>"
                End If
            End If
            i = i + 1
        Loop Until i > iTotal
        If iDummyCount > iDetectedCount Then
            For j = i To UBound(arrDummy)
                cSplit = Split(arrDummy(j), "|")
                '   Don't write if a weekend
                If Weekday(cSplit(0)) <> 7 And Weekday(cSplit(0)) <> 1 Then
                    If IsEven(i) = True Then
                        myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
                    Else
                        myFile.WriteLine ("<tr bgcolor ='ivory'>")
                    End If
                    '   Write a row to the table
                    myFile.WriteLine "<td>" & cSplit(0) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
                End If
            Next
        End If
        '   Write the remainder of the html
        myFile.WriteLine "</table></body></html>"
        '   Close the file
        myFile.Close
    
        '   Now put the data into the html file
        '  Open the template file
        Set mySource = fso.OpenTextFile(strTemplatePath & "NewDefectDailyTemplate.txt", ForReading)
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDaily.aspx", ForWriting, True)
        Do While mySource.AtEndOfStream <> True
            rc = mySource.ReadLine
            If InStr(1, rc, "AcceptedDefectsPoints") > 0 Then
                '   Loop round our total array
                For i = 0 To UBound(arrAccepted)
                    '   Split the file
                    mySplit = Split(arrAccepted(i), "|")
                    If mySplit(0) = "Project Start" Then
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                    Else
                        If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                            '   Re-format date value
                            myDate = Format(mySplit(0), "dd mmm yy")
                            '   Write the value
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                        End If
                    End If
                Next
            Else
                If InStr(1, rc, "FixedDefectsPoints") > 0 Then
                    '   Loop round our Fixed array
                    For i = 0 To UBound(arrFixed)
                        '   Split the file
                        mySplit = Split(arrFixed(i), "|")
                        If mySplit(0) = "Project Start" Then
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                        Else
                            If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                '   Re-format date value
                                myDate = Format(mySplit(0), "dd mmm yy")
                                '   Write the value
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                            End If
                        End If
                    Next
                Else
                    If InStr(1, rc, "TestedDefectsPoints") > 0 Then
                        '   Loop round our Tested array
                        For i = 0 To UBound(arrTested)
                            '   Split the file
                            mySplit = Split(arrTested(i), "|")
                            If mySplit(0) = "Project Start" Then
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                            Else
                                If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                    '   Re-format date value
                                    myDate = Format(mySplit(0), "dd mmm yy")
                                    '   Write the value
                                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                                End If
                            End If
                        Next
                    Else
                        If InStr(1, rc, "HiddenPoints") > 0 Then
                            '   Loop round our Dummy array
                            For i = 0 To UBound(arrDummy)
                                '   Split the file
                                mySplit = Split(arrDummy(i), "|")
                                If mySplit(0) = "Project Start" Then
                                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                                Else
                                    '   Don't write if on a weekend
                                    If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                        '   Re-format date value
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
        Loop
        
        '   Close the files
        myDest.Close
        mySource.Close
        
        '   Now the detected vs the closed
        Set mySource = fso.OpenTextFile(strTemplatePath & "DefectsDetectedvsClosedTemplate.txt", ForReading)
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDetectedvsClosed.aspx", ForWriting, True)
        Do While mySource.AtEndOfStream <> True
            rc = mySource.ReadLine
            If InStr(1, rc, "DetectedPoints") > 0 Then
                '   Loop round our total array
                For i = 0 To UBound(arrTotal)
                    '   Split the file
                    mySplit = Split(arrTotal(i), "|")
                    If mySplit(0) = "Project Start" Then
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                    Else
                        If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                            '   Re-format date value
                            myDate = Format(mySplit(0), "dd mmm yy")
                            '   Write the value
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                        End If
                    End If
                Next
            Else
                If InStr(1, rc, "ClosedPoints") > 0 Then
                    '   Loop round our Fixed array
                    For i = 0 To UBound(arrClosed)
                        '   Split the file
                        mySplit = Split(arrClosed(i), "|")
                        If mySplit(0) = "Project Start" Then
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                        Else
                            If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                '   Re-format date value
                                myDate = Format(mySplit(0), "dd mmm yy")
                                '   Write the value
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                            End If
                        End If
                    Next
                Else
                    If InStr(1, rc, "HiddenPoints") > 0 Then
                            '   Loop round our Dummy array
                            For i = 0 To UBound(arrDummy)
                                '   Split the file
                                mySplit = Split(arrDummy(i), "|")
                                If mySplit(0) = "Project Start" Then
                                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(0) & Chr(34) & " />"
                                Else
                                    '   Don't write if a weekend
                                    If Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1 Then
                                        '   Re-format date value
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
        Loop
        
        '   Close the files
        myDest.Close
        mySource.Close
        
        '   Now write out the Detected vs Closed table info
        fso.CopyFile strTemplatePath & "DefectsDetectedvsClosedTableTemplate.txt", strFolderPath & strPathandFileName & "-DefectsDetectedvsClosedTable.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsDetectedvsClosedTable.asp", ForAppending, True)
        
        '   Write out the table header details etc
        myFile.WriteLine "<table border=1 align=center>"
        myFile.WriteLine "<tr><th colspan=7 align=center>" & strHeader & "</th></tr>"
        myFile.WriteLine "<tr bgcolor='#CCCCCC'><td>&nbsp;</td><td>Detected</td><td>Closed</td></tr>"
        iDetectedCount = UBound(arrTotal)
        iDummyCount = UBound(arrDummy)
        If iDetectedCount = iDummyCount Then
            iTotal = iDummyCount
        Else
            iTotal = iDetectedCount
        End If
        i = 0
        Do
            '   Get the values from each of the arrays for this array element
            aSplit = Split(arrTotal(i), "|")
            bSplit = Split(arrClosed(i), "|")
            cSplit = Split(arrDummy(i), "|")
            
            '   If we're on the first element then just default to project start and zeros
            If i = 0 Then
                myFile.WriteLine "<tr bgcolor ='#B5EAAA'><td>Project Start</td><td>0</td><td>0</td></tr>"
            Else
                If Weekday(aSplit(0)) <> 7 And Weekday(aSplit(0)) <> 1 Then
                    If IsEven(i) = True Then
                        myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
                    Else
                        myFile.WriteLine ("<tr bgcolor ='ivory'>")
                    End If
                    '   Write a row to the table
                    myFile.WriteLine "<td>" & aSplit(0) & "</td><td>" & aSplit(1) & "</td><td>" & bSplit(1) & "</td></tr>"
                End If
            End If
            i = i + 1
        Loop Until i > iTotal
        If iDummyCount > iDetectedCount Then
            For j = i To UBound(arrDummy)
                cSplit = Split(arrDummy(j), "|")
                '   Don't write if a weekend
                If Weekday(cSplit(0)) <> 7 And Weekday(cSplit(0)) <> 1 Then
                    If IsEven(i) = True Then
                        myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
                    Else
                        myFile.WriteLine ("<tr bgcolor ='ivory'>")
                    End If
                    '   Write a row to the table
                    myFile.WriteLine "<td>" & cSplit(0) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
                End If
            Next
        End If
        '   Write the remainder of the html
        myFile.WriteLine "</table></body></html>"
        '   Close the file
        myFile.Close
        
        '   Update the defect status file
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForReading, True)
        strText = myDest.ReadAll
        myDest.Close
        
        '   Update the graph links
        strText = Replace(strText, "DetectedvsClosed", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetectedvsClosed.aspx" & Chr(34))
        
        '   Write it back out
        Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForWriting, True)
        myDest.WriteLine strText
        myDest.Close
    Else
        '   Copy the defect error page to the test runs page
        fso.CopyFile strTemplatePath & "DefectsMissingTemplate.txt.txt", strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt"
        '   Open the file and change the header
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        '   Replace the text
        strText = Replace(strText, "strHeader", "No Defect Details for " & strHeader)
        
        '   Write it back out
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForWriting, True)
        myFile.WriteLine strText
        myFile.Close
        
        '   Set the flag
        blnDontRun = True
            
    End If
    
Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function CountDefectsByDate(ByVal strType As String, ByVal strDateFilter As String, Optional strFixedOnDate As String, Optional strTestedOnDate As String) As Integer
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
Dim iCount As Integer
'Function to count the number of defects returned using a filter
'Always filter by project

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    '   Set filter
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    
    tdcBugFilter.Filter("BG_PROJECT") = "'" & strProjectName & "'"
    tdcBugFilter.Filter(strTestPhaseBugLabel) = "'" & strTestPhase & "'"
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    Select Case strType
        Case "Accepted", "New"
            tdcBugFilter.Filter("BG_DETECTION_DATE") = strDateFilter
        Case "Fixed"
            tdcBugFilter.Filter(strFixedOnDateLabel) = strDateFilter
        Case "Tested"
            tdcBugFilter.Filter(strTestedOnDateLabel) = strDateFilter
        Case "Closed"
            tdcBugFilter.Filter(strClosedOnDateLabel) = strDateFilter
    End Select
    
    Set colBugList = tdcBugFilter.NewList
    
    '   If we're doing accepted then use the count minus Rejected, Duplicate and On Hold status
    '   If the status is Closed then check if the root cause is Duplicate or Rejected and reduce count
    If strType = "Accepted" Then
        iCount = colBugList.Count
        For Each TheBug In colBugList
            Select Case TheBug.Status
                Case "Rejected", "Duplicate", "On Hold"
                    iCount = iCount - 1
                Case "Closed"
                    rc = TheBug.Field(strRootCauseLabel)
                    If rc = "Rejected" Or rc = "Duplicate" Then
                        iCount = iCount - 1
                    End If
            End Select
        Next
        CountDefectsByDate = iCount
    Else
        '   If the type is tested, then check if the status is 'Failed Testing' and remove from count
        If strType = "Tested" Then
            iCount = colBugList.Count
            For Each TheBug In colBugList
               Select Case TheBug.Status
                  Case "Failed Testing"
                     iCount = iCount - 1
               End Select
            Next
            CountDefectsByDate = iCount
        Else
            CountDefectsByDate = colBugList.Count
        End If
    End If
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function GetDefectsbyHistory()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
Dim hst As TDAPIOLELib.History
Dim hstRec As TDAPIOLELib.HistoryRecord
Dim hstList As TDAPIOLELib.list
Dim iCount
Dim iNewCount As Integer, iOpenCount As Integer, iAssignedCount As Integer
Dim myDateFilter As String
Dim rc
Dim myArr()
Dim dateStartDate As Date, dateEndDate As Date
intClosed = 0
intDupRej = 0
intFixed = 0
intNew = 0
intOnHold = 0
intOpen = 0
intReopen = 0
intTested = 0

    
    '   Set up the main filter
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    
    Set colBugList = tdcBugFilter.NewList

    '   See if we've got any bugs
    iNoItems = colBugList.Count
    If iNoItems = 0 Then
        '  Open the defects by history template
        Set myFile = fso.OpenTextFile(strTemplatePath & "DefectsByHistoryTemplate.txt", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        '   Replace the title data
        strNewText = Replace(strText, "ProjectTitle", strPathandFileName)
        '   Write the new text to the asp file
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
        '   Open the file again, for appending to
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForAppending, True)
        '   Write out the tables
        myFile.WriteLine ("<html><body>")
        myFile.WriteLine ("<table width='700' align='center' border='1'>")
        myFile.WriteLine ("<tr><th colspan=14>" & strHeader & "</th></tr>")
        myFile.WriteLine ("<tr bgcolor='#CCCCCC'><th><font size='-1'>Date</font></th><th><font size='-1'>New</font></th><th><font size='-1'>Assigned</font></th>")
        myFile.WriteLine ("<th><font size='-1'>Open</font></th><th><font size='-1'>Fixed</font></th><th><font size='-1'>Ready</br>For</br> Testing</font></th>")
        myFile.WriteLine ("<th><font size='-1'>Failed</br>Testing</font></th><th><font size='-1'>Tested</font></th><th><font size='-1'>Reopen</font></th>")
        myFile.WriteLine ("<th><font size='-1'>Duplicate</font></th><th><font size='-1'>Rejected</font></th><th><font size='-1'>On</br>Hold</font></th>")
        myFile.WriteLine ("<th><font size='-1'>Closed</font></th><th><font size='-1'>Total</font></th></tr>")
        myFile.WriteLine ("</table>")
        myFile.WriteLine ("<%FinishPage();%>")
        myFile.WriteLine ("</body></html>")
        myFile.Close
        Exit Function
    End If

    '   Get the 1st bug bought back and get its 1st status date
    date1stDate = FindFirstBugDate

    '   Get the dates and statuses
    myDates = GetDatesAndStatus(date1stDate, TodaysDate)

    '   Get the number of items in the bug list
    iNoItems = colBugList.Count
    '   Loop round the bug items
    For i = 1 To iNoItems
        Set bg = colBugList.Item(i)
        Set hst = bg.History
        Set hstList = hst.NewList("")

        If hstList.Count > 0 Then
            iCount = -1
            ' Loop round the history items for this bug
            For j = hstList.Count To 1 Step -1
                Set hstRec = hstList.Item(j)
                '   If we find a status field
                If hstRec.FieldName = "BG_STATUS" Then
                    '   Get the date
                    dateThisDate = Left(hstRec.ChangeDate, 10)
                    '   Get the status
                    strStatus = hstRec.NewValue
                    '   Add to array if <= TodaysDate
                    If dateThisDate <= TodaysDate Then
                        '  See if the previous status is any of the 'dodgy' ones
                        Select Case (strStatus)
                            Case "In Progress"
                                strStatus = "Open"
                            Case "Ready for Testing"
                                strStatus = "Ready For Testing"
                            Case "Status_Closed"
                                strStatus = "Closed"
                        End Select
                        '   Add to the array
                        iCount = iCount + 1
                        ReDim Preserve myArr(iCount)
                        myArr(iCount) = dateThisDate & "|" & strStatus
                    End If
                End If
            Next
            '   Break out to next record if there's no status history
            If iCount = -1 Then
                GoTo NextRecord
            End If
            '   Work around the array putting values into correct places in the date array
            For iEle = 0 To UBound(myArr)
                mySplit = Split(myArr(iEle), "|")
                
                '   Find this date
                Select Case mySplit(1)
                    Case "New"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Assigned"
                        If UBound(myArr) > 0 Then
                            '   See what it's status was before this one
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Open", "In Progress"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Fixed"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Ready For Testing", "Ready for Testing"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Failed Testing"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Tested"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Reopen"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Duplicate"
                        '   Make sure we've got more than 1 in the array
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Rejected"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "On Hold"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case "Closed", "Status_Closed"
                        If UBound(myArr) > 0 Then
                            If iEle <> 0 Then
                                strPrevStatus = Mid(myArr(iEle - 1), 12)
                                '  See if the previous status is any of the 'dodgy' ones
                                Select Case strPrevStatus
                                    Case "In Progress"
                                        strPrevStatus = "Open"
                                    Case "Ready for Testing"
                                        strPrevStatus = "Ready For Testing"
                                    Case "Status_Closed"
                                        strPrevStatus = "Closed"
                                End Select
                                strThisDate = Left(myArr(iEle), 10)
                                '   Remove 1 from this status count on the day
                                rc = objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus)
                                rc = rc - 1
                                objDatesStatusDictionary.Item(strThisDate & "|" & strPrevStatus) = rc
                            End If
                        End If
                        '   Now add 1 to our column
                        rc = objDatesStatusDictionary.Item(myArr(iEle))
                        rc = rc + 1
                        objDatesStatusDictionary.Item(myArr(iEle)) = rc
                    Case Else
                        MsgBox mySplit(1)
                End Select
            Next
        End If
NextRecord:
    Next
    
    '   Now accumlate all the values
    AccumulateData
    
    '  Open the defects by history template
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectsByHistoryTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    '   Replace the title data
    strNewText = Replace(strText, "ProjectTitle", strPathandFileName)
    '   Write the new text to the asp file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForWriting, True)
    myFile.WriteLine strNewText
    myFile.Close
    '   Open the file again, for appending to
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForAppending, True)
    '   Write out the tables
    myFile.WriteLine ("<table width='700' align='center' border='1'>")
    myFile.WriteLine ("<tr><th colspan=14>" & strHeader & "</th></tr>")
    myFile.WriteLine ("<tr bgcolor='#CCCCCC'><th><font size='-1'>Date</font></th><th><font size='-1'>New</font></th><th><font size='-1'>Assigned</font></th>")
    myFile.WriteLine ("<th><font size='-1'>Open</font></th><th><font size='-1'>Fixed</font></th><th><font size='-1'>Ready</br>For</br> Testing</font></th>")
    myFile.WriteLine ("<th><font size='-1'>Failed</br>Testing</font></th><th><font size='-1'>Tested</font></th><th><font size='-1'>Reopen</font></th>")
    myFile.WriteLine ("<th><font size='-1'>Duplicate</font></th><th><font size='-1'>Rejected</font></th><th><font size='-1'>On</br>Hold</font></th>")
    myFile.WriteLine ("<th><font size='-1'>Closed</font></th><th><font size='-1'>Total</font></th></tr>")
    
    '  Put date dictionary into array
    myKeys = objDateDictionary.Keys
    
    '  Loop round arrays putting values into correct slots
    For i = UBound(myKeys) To 0 Step -1
        '  See if we colour this row
        If IsEven(i) = True Then
            myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
        Else
            myFile.WriteLine ("<tr bgcolor ='ivory'>")
        End If
        '  Write the date to the table
        myFile.WriteLine ("<td><font size='-1'>" & myKeys(i) & "</font></td>")
        '  Get the values for this date
        iNew = objDatesStatusDictionary.Item(myKeys(i) & "|" & "New")
        iAssigned = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Assigned")
        iOpen = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Open")
        iFixed = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Fixed")
        iReady = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Ready For Testing")
        iFailed = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Failed Testing")
        iTested = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Tested")
        iReopen = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Reopen")
        iDuplicate = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Duplicate")
        iRejected = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Rejected")
        iOnHold = objDatesStatusDictionary.Item(myKeys(i) & "|" & "On Hold")
        iClosed = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Closed")

        '  Add them to get total
        iTotal = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed

        '  Write them out to the table
        myFile.WriteLine ("<td align='center'><font size='-1'>" & iNew & "</font></td><td align='center'><font size='-1'>" & iAssigned & "</font></td><td align='center'><font size='-1'>" & iOpen & "</font></td><td align='center'><font size='-1'>" & iFixed & "</font></td><td align='center'><font size='-1'>" & iReady & "</font></td>")
        myFile.WriteLine ("<td align='center'><font size='-1'>" & iFailed & "</font></td><td align='center'><font size='-1'>" & iTested & "</font></td><td align='center'><font size='-1'>" & iReopen & "</font></td><td align='center'><font size='-1'>" & iDuplicate & "</font></td>")
        myFile.WriteLine ("<td align='center'><font size='-1'>" & iRejected & "</font></td><td align='center'><font size='-1'>" & iOnHold & "</font></td><td align='center'><font size='-1'>" & iClosed & "</font></td><td align='center'><font size='-1'>" & iTotal & "</font></td>")
        myFile.WriteLine ("</tr>")
    Next

    myFile.WriteLine ("</table>")
    myFile.WriteLine ("<%FinishPage();%>")
    myFile.WriteLine ("</body>")
    myFile.WriteLine ("</html>")
    myFile.Close
    
End Function
Public Function GetDefectsOverTime()
Dim myArr()
Dim myDummy()
iCount = -1
iDum = -1

    '  Put date dictionary into array
    myKeys = objDateDictionary.Keys
    
    '   Get the key dates for our graph
    dateStartDate = CDate(myKeys(LBound(myKeys)))
    dateThisEndDate = myKeys(UBound(myKeys))

    '   Get the date for the start of the daily dates
    dateLastWeek = DateAdd("d", -7, TodaysDate)
    '   See if this date is before our start date
    If dateLastWeek > dateStartDate Then
        '   Now go back a month from the day start
        datePreviousMonth = DateAdd("m", -1, dateLastWeek)
        If Weekday(datePreviousMonth) = 1 Then
            datePreviousMonth = DateAdd("d", 1, datePreviousMonth)
        Else
            If Weekday(datePreviousMonth) = 7 Then
                datePreviousMonth = DateAdd("d", 2, datePreviousMonth)
            End If
        End If
        '   See if this date is before our start date
        If datePreviousMonth < dateStartDate Then
            arrWeekends = FindWeekends(dateStartDate, dateLastWeek)
            arrMonths = FindMonths(dateStartDate, datePreviousMonth)
        Else
            '   Get the previous months
            arrMonths = FindMonths(dateStartDate, datePreviousMonth)
            If arrMonths(0) <> "No Month" Then
                arrWeekends = FindWeekends(arrMonths(UBound(arrMonths)), dateLastWeek)
            Else
                If datePreviousMonth > strSDate Then
                    arrWeekends = FindWeekends(datePreviousMonth, dateLastWeek)
                Else
                    arrWeekends = FindWeekends(dateStartDate, dateLastWeek)
                End If
            End If
        End If
    
        '   See if we're just doing day, weeks and days or months, weeks and days
        If arrMonths(0) = "No Month" And arrWeekends(0) = "No Weeks" Then
            '  Loop round arrays putting values into correct slots
            For i = 0 To UBound(myKeys)
            
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(myKeys(i) & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(myKeys(i) & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(myKeys(i) & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
                
                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = myKeys(i) & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
            
            Next
            blnGoNoFurther = True
            GoTo WriteHTML
        End If
        If arrMonths(0) = "No Month" And arrWeekends(0) <> "No Weeks" Then
            
            '   Add the start date values as a starting point
            iOutstanding = 0
            iTotalFixed = 0
            iTotalTested = 0
            
            '  Get the values for this date
            iNew = objDatesStatusDictionary.Item(dateStartDate & "|" & "New")
            iAssigned = objDatesStatusDictionary.Item(dateStartDate & "|" & "Assigned")
            iOpen = objDatesStatusDictionary.Item(dateStartDate & "|" & "Open")
            iFixed = objDatesStatusDictionary.Item(dateStartDate & "|" & "Fixed")
            iReady = objDatesStatusDictionary.Item(dateStartDate & "|" & "Ready For Testing")
            iFailed = objDatesStatusDictionary.Item(dateStartDate & "|" & "Failed Testing")
            iTested = objDatesStatusDictionary.Item(dateStartDate & "|" & "Tested")
            iReopen = objDatesStatusDictionary.Item(dateStartDate & "|" & "Reopen")
            iDuplicate = objDatesStatusDictionary.Item(dateStartDate & "|" & "Duplicate")
            iRejected = objDatesStatusDictionary.Item(dateStartDate & "|" & "Rejected")
            iOnHold = objDatesStatusDictionary.Item(dateStartDate & "|" & "On Hold")
            iClosed = objDatesStatusDictionary.Item(dateStartDate & "|" & "Closed")

            '  Add them to get Oustanding
            iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
            '   Add them to get Accepted
            iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
            '   Add then to get Total Tested
            iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
                
            '   Add to array
            iCount = iCount + 1
            ReDim Preserve myArr(iCount)
            myArr(iCount) = dateStartDate & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
            
            For i = 0 To UBound(arrWeekends)
                iOutstanding = 0
                iTotalFixed = 0
                iTotalTested = 0
                
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
   
                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = arrWeekends(i) & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
            Next
    
            '   Now do days
            iOutstanding = 0
            iAccepted = 0
            iDetected = 0
            ThisDate = DateAdd("d", 1, arrWeekends(UBound(arrWeekends)))
            Do
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(ThisDate & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(ThisDate & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(ThisDate & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(ThisDate & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(ThisDate & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(ThisDate & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(ThisDate & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(ThisDate & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(ThisDate & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(ThisDate & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(ThisDate & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(ThisDate & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
   
                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = ThisDate & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
               
                '   Up the date
                ThisDate = DateAdd("d", 1, ThisDate)
                
            Loop Until ThisDate > dateThisEndDate
        Else
            '   Add the start date values as a starting point
            iOutstanding = 0
            iAccepted = 0
            iDetected = 0
            
            '  Get the values for this date
            iNew = objDatesStatusDictionary.Item(dateStartDate & "|" & "New")
            iAssigned = objDatesStatusDictionary.Item(dateStartDate & "|" & "Assigned")
            iOpen = objDatesStatusDictionary.Item(dateStartDate & "|" & "Open")
            iFixed = objDatesStatusDictionary.Item(dateStartDate & "|" & "Fixed")
            iReady = objDatesStatusDictionary.Item(dateStartDate & "|" & "Ready For Testing")
            iFailed = objDatesStatusDictionary.Item(dateStartDate & "|" & "Failed Testing")
            iTested = objDatesStatusDictionary.Item(dateStartDate & "|" & "Tested")
            iReopen = objDatesStatusDictionary.Item(dateStartDate & "|" & "Reopen")
            iDuplicate = objDatesStatusDictionary.Item(dateStartDate & "|" & "Duplicate")
            iRejected = objDatesStatusDictionary.Item(dateStartDate & "|" & "Rejected")
            iOnHold = objDatesStatusDictionary.Item(dateStartDate & "|" & "On Hold")
            iClosed = objDatesStatusDictionary.Item(dateStartDate & "|" & "Closed")

            '  Add them to get Oustanding
            iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
            '   Add them to get Accepted
            iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
            '   Add then to get Total Tested
            iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
 
            '   Add to array
            iCount = iCount + 1
            ReDim Preserve myArr(iCount)
            myArr(iCount) = dateStartDate & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected

            For i = 0 To UBound(arrMonths)
                iOutstanding = 0
                iAccepted = 0
                iDetected = 0
                
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(arrMonths(i) & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed

                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = arrMonths(i) & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
            Next
            
            '   Now do weeks
            
            '  Loop round arrays putting values into correct slots
            For i = 0 To UBound(arrWeekends)
                iOutstanding = 0
                iAccepted = 0
                iDetected = 0
                
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(arrWeekends(i) & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
    
                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = arrWeekends(i) & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
            Next
            
            '   Now do days
            iOutstanding = 0
            iAccepted = 0
            iDetected = 0
            ThisDate = DateAdd("d", 1, arrWeekends(UBound(arrWeekends)))
            Do
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(ThisDate & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(ThisDate & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(ThisDate & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(ThisDate & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(ThisDate & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(ThisDate & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(ThisDate & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(ThisDate & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(ThisDate & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(ThisDate & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(ThisDate & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(ThisDate & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
                    
                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = ThisDate & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
               
                '   Up the date
                ThisDate = DateAdd("d", 1, ThisDate)
                
            Loop Until ThisDate > dateThisEndDate
            
        End If
    Else
            iOutstanding = 0
            iAccepted = 0
            iDetected = 0
            ThisDate = dateStartDate
            Do
                '  Get the values for this date
                iNew = objDatesStatusDictionary.Item(ThisDate & "|" & "New")
                iAssigned = objDatesStatusDictionary.Item(ThisDate & "|" & "Assigned")
                iOpen = objDatesStatusDictionary.Item(ThisDate & "|" & "Open")
                iFixed = objDatesStatusDictionary.Item(ThisDate & "|" & "Fixed")
                iReady = objDatesStatusDictionary.Item(ThisDate & "|" & "Ready For Testing")
                iFailed = objDatesStatusDictionary.Item(ThisDate & "|" & "Failed Testing")
                iTested = objDatesStatusDictionary.Item(ThisDate & "|" & "Tested")
                iReopen = objDatesStatusDictionary.Item(ThisDate & "|" & "Reopen")
                iDuplicate = objDatesStatusDictionary.Item(ThisDate & "|" & "Duplicate")
                iRejected = objDatesStatusDictionary.Item(ThisDate & "|" & "Rejected")
                iOnHold = objDatesStatusDictionary.Item(ThisDate & "|" & "On Hold")
                iClosed = objDatesStatusDictionary.Item(ThisDate & "|" & "Closed")
    
                '  Add them to get Oustanding
                iOutstanding = iNew + iAssigned + iOpen + iReopen + iFailed
                '   Add them to get Accepted
                iAccepted = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iClosed
                '   Add then to get Total Tested
                iDetected = iNew + iAssigned + iOpen + iFixed + iReady + iFailed + iTested + iReopen + iDuplicate + iRejected + iOnHold + iClosed
                    
                '   Add to array
                iCount = iCount + 1
                ReDim Preserve myArr(iCount)
                myArr(iCount) = ThisDate & "|" & iOutstanding & "|" & iAccepted & "|" & iDetected
               
                '   Up the date
                ThisDate = DateAdd("d", 1, ThisDate)
                
            Loop Until ThisDate > dateThisEndDate
    End If
    
    '   Now get just the planned values out into the future if required
    If dateEndDate = "00:00:00" Then
        blnGoNoFurther = True
        GoTo WriteHTML
    Else
        blnGoNoFurther = False
    End If

    '   See if we've got any test sets and so an end date
    If blnNoTestSets = True Then
        dateEndDate = TodaysDate
    End If
    If dateEndDate > TodaysDate Then
        iDays = DateDiff("d", TodaysDate, dateEndDate)
        If iDays <= 30 Then
            strSDate = DateAdd("d", 1, TodaysDate)
            Do While strSDate <= dateEndDate
                iDum = iDum + 1
                ReDim Preserve myDummy(iDum)
                myDummy(iDum) = strSDate & "|0"
                strSDate = strSDate + 1
            Loop
            GoTo WriteHTML
        End If
        strSDate = DateAdd("d", 1, TodaysDate)
        '   Get the date next week
        dateNextWeek = DateAdd("d", 7, TodaysDate)
        '   See if it's past our end date
        If dateNextWeek >= dateEndDate Then
            '   Get the last value reported cos we're just going to repeat this
            Do While strSDate <= dateEndDate
                iDum = iDum + 1
                ReDim Preserve myDummy(iDum)
                myDummy(iDum) = strSDate & "|0"
                strSDate = strSDate + 1
            Loop
            GoTo WriteHTML
        End If
        dateNextMonth = DateAdd("m", 1, dateNextWeek)
        '   See if it's past our end date
        If dateNextMonth >= dateEndDate Then
            arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
            If arrWeekends(0) <> "No Weeks" Then
                If arrWeekends(0) < dateEndDate Then
                    Do While strSDate <= dateNextWeek
                        iDum = iDum + 1
                        ReDim Preserve myDummy(iDum)
                        myDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                    For Each Ele In arrWeekends
                        iDum = iDum + 1
                        ReDim Preserve myDummy(iDum)
                        myDummy(iDum) = Ele & "|0"
                        strSDate = Ele
                    Next
                    strSDate = strSDate + 1
                    Do While strSDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve myDummy(iDum)
                        myDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                Else
                    Do While strSDate <= dateEndDate
                        iDum = iDum + 1
                        ReDim Preserve myDummy(iDum)
                        myDummy(iDum) = strSDate & "|0"
                        strSDate = strSDate + 1
                    Loop
                End If
            Else
                Do While strSDate <= dateEndDate
                    iDum = iDum + 1
                    ReDim Preserve myDummy(iDum)
                    myDummy(iDum) = strSDate & "|0"
                    strSDate = strSDate + 1
                Loop
                GoTo WriteHTML
            End If
        End If
        '   See how many days we're dealing with
        iDays = DateDiff("d", dateNextMonth, dateEndDate)
            If iDays < 0 Then
                GoTo WriteHTML
            End If
            If iDays <= 30 Then
                arrWeekends = FindWeekends(dateNextWeek, dateEndDate)
                '   Write out the first weeks worth of days
                Do While strSDate <= dateNextWeek
                    iDum = iDum + 1
                    ReDim Preserve myDummy(iDum)
                    myDummy(iDum) = strSDate & "|0"
                    strSDate = strSDate + 1
                Loop
                '   Now write out the remaining weeks
                For Each Ele In arrWeekends
                    iDum = iDum + 1
                    ReDim Preserve myDummy(iDum)
                    myDummy(iDum) = Ele & "|0"
                    myDate = Ele
                Next
            Else
                arrMonths = FindMonths(dateNextMonth, dateEndDate)
                arrWeekends = FindWeekends(dateNextWeek, dateNextMonth)
                Do While strSDate <= dateNextWeek
                    iDum = iDum + 1
                    ReDim Preserve myDummy(iDum)
                    myDummy(iDum) = strSDate & "|0"
                    strSDate = strSDate + 1
                Loop
                For Each Ele In arrWeekends
                    iDum = iDum + 1
                    ReDim Preserve myDummy(iDum)
                    myDummy(iDum) = Ele & "|0"
                    myDate = Ele
                Next
                For Each Ele In arrMonths
                    iDum = iDum + 1
                    ReDim Preserve myDummy(iDum)
                    myDummy(iDum) = Ele & "|0"
                    myDate = Ele
                Next
            End If
        
    Else
        blnGoNoFurther = True
    End If
    
WriteHTML:

    '   Open the template
    Set mySource = fso.OpenTextFile(strTemplatePath & "NewDefectsOverTimeTemplate.txt", ForReading)
    '   Open the output file
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsOverTime.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "OutstandingPoints") > 0 Then
            '   Write the project start value
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & "0" & Chr(34) & " AxisLabel=" & Chr(34) & "Project Start" & Chr(34) & " />"
            For Each Ele In myArr
                '   Split the file
                mySplit = Split(Ele, "|")
                '   Drop if a weekend with no data
                If (Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1) Then
                   '   Re-format the date part
                   myDate = Format(mySplit(0), "dd mmm yy")
                   '   Write the value
                   myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                End If
           Next
        Else
            If InStr(1, rc, "AcceptedPoints") > 0 Then
                '   Write the project start value
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & "0" & Chr(34) & " AxisLabel=" & Chr(34) & "Project Start" & Chr(34) & " />"
                For Each Ele In myArr
                    '   Split the file
                    mySplit = Split(Ele, "|")
                    '   Drop if a weekend with no data
                    If (Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1) Then
                        '   Re-format the date part
                        myDate = Format(mySplit(0), "dd mmm yy")
                        '   Write the value
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                    End If
                Next
            Else
                If InStr(1, rc, "HiddenPoints") > 0 Then
                    '   Write the project start value
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & "0" & Chr(34) & " AxisLabel=" & Chr(34) & "Project Start" & Chr(34) & " />"
                    For Each Ele In myArr
                        '   Split the file
                        mySplit = Split(Ele, "|")
                        '   Drop if a weekend with no data
                        If (Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1) Then
                            '   Re-format the date part
                            myDate = Format(mySplit(0), "dd mmm yy")
                            '   Write the value
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & "0" & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                        End If
                    Next
                    If blnGoNoFurther = False Then
                        For Each Ele In myDummy
                            '   Split the file
                            mySplit = Split(Ele, "|")
                            If (Weekday(mySplit(0)) <> 7 And Weekday(mySplit(0)) <> 1) Then
                                '  Re-format the date part
                                myDate = Format(mySplit(0), "dd mmm yy")
                                '   Write the value
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & "0" & Chr(34) & " AxisLabel=" & Chr(34) & myDate & Chr(34) & " />"
                            End If
                        Next
                    End If
                Else
                    myDest.WriteLine rc
                End If
            End If
        End If
    Loop
    
    '   Close files
    mySource.Close
    myDest.Close

    '   Now write out the Detects over time table info
    fso.CopyFile strTemplatePath & "DefectsOverTimeTableTemplate.txt", strFolderPath & strPathandFileName & "-DefectsOverTimeTable.asp"
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsOverTimeTable.asp", ForAppending, True)
    
    '   Write out the table header details etc
    myFile.WriteLine "<table border=1 align=center>"
    myFile.WriteLine "<tr><th colspan=7 align=center>" & strHeader & "</th></tr>"
    myFile.WriteLine "<tr bgcolor='#CCCCCC'><td>&nbsp;</td><td>Accepted</td><td>Outstanding</td></tr>"
    iCount = UBound(myArr)
    If blnGoNoFurther = False Then
        iDummyCount = UBound(myDummy)
        If iCount = iDummyCount Then
            iTotal = iDummyCount
        Else
            iTotal = iCount
        End If
    Else
        iTotal = iCount
    End If
    '   Write the start line
    myFile.WriteLine "<tr bgcolor ='#B5EAAA'><td>Project Start</td><td>0</td><td>0</td></tr>"
    i = 0
    Do
        '   Get the values from each of the arrays for this array element
        aSplit = Split(myArr(i), "|")
        If Weekday(aSplit(0)) <> 7 And Weekday(aSplit(0)) <> 1 Then
            '   If we're on the first element then just default to project start and zeros
            If IsEven(i) = True Then
                myFile.WriteLine ("<tr bgcolor ='ivory'>")
            Else
                myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
            End If
            '   Write a row to the table
            myFile.WriteLine "<td>" & aSplit(0) & "</td><td>" & aSplit(2) & "</td><td>" & aSplit(1) & "</td></tr>"
        End If
        i = i + 1
    Loop Until i > iTotal
    If blnGoNoFurther = False Then
        For j = 0 To UBound(myDummy)
            cSplit = Split(myDummy(j), "|")
            If Weekday(cSplit(0)) <> 7 And Weekday(cSplit(0)) <> 1 Then
                If IsEven(i) = True Then
                    myFile.WriteLine ("<tr bgcolor ='ivory'>")
                Else
                    myFile.WriteLine ("<tr bgcolor ='#B5EAAA'>")
                End If
                '   Write a row to the table
                myFile.WriteLine "<td>" & cSplit(0) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
            End If
        Next
    End If
    
    '   Write the remainder of the html
    myFile.WriteLine "</table></body></html>"
    '   Close the file
    myFile.Close
    
    '   Update the defect status file
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForReading, True)
    strText = myDest.ReadAll
    myDest.Close
    
    '   Update the graph link
    strText = Replace(strText, "DefectsOverTime", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsOverTime.aspx" & Chr(34))
    
    '   Write it back out
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", ForWriting, True)
    myDest.WriteLine strText
    myDest.Close
    
End Function
Public Function GetDefectAgeBySeverity()

    ' Count the new - fixed defects by age and add them to the data sheet.
    arrNewFixed = DefectAgeBySeverity("New-Fixed")
    arrFixedTested = DefectAgeBySeverity("Fixed-Tested")
    arrTimeOpen = DefectAgeBySeverity("TimeOpen")
    
    '   Build the html for new-fixed
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "DefectNewFixedTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectNewFixed.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "1-CriticalPoints") > 0 Then
            '   Split the file
            mySplit = Split(arrNewFixed(0), "|")
                '   Write the value
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
        Else
            If InStr(1, rc, "2-HighPoints") > 0 Then
                '   Split the file
                mySplit = Split(arrNewFixed(1), "|")
                '   Write the value
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
            Else
                If InStr(1, rc, "3-MediumPoints") > 0 Then
                    '   Split the file
                    mySplit = Split(arrNewFixed(2), "|")
                    '   Write the value
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
                Else
                    If InStr(1, rc, "4-LowPoints") > 0 Then
                        '   Split the file
                        mySplit = Split(arrNewFixed(3), "|")
                        '   Write the value
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
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
    
    '   See if we need to amend or remove the y axis interval
    iCount = 0
    mySplit = Split(arrNewFixed(0))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrNewFixed(1))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            iCount = iCount + CInt(NextEle)
        Next
    Next
    mySplit = Split(arrNewFixed(2))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            iCount = iCount + CInt(NextEle)
        Next
    Next
    mySplit = Split(arrNewFixed(3))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            iCount = iCount + CInt(NextEle)
        Next
    Next
    
    '   Open the new - fixed file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectNewFixed.aspx", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    '   Look at the count to decide how we change the y axis interval
    Select Case iCount
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
            strText = Replace(strText, "ReplaceYAxis", "Interval=" & Chr(34) & "1" & Chr(34))
        Case Else
            strText = Replace(strText, "ReplaceYAxis", "")
    End Select
    '   Replace in the file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectNewFixed.aspx", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "DefectFixedTestedTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFixedTested.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "1-CriticalPoints") > 0 Then
            '   Split the file
            mySplit = Split(arrFixedTested(0), "|")
                '   Write the value
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
        Else
            If InStr(1, rc, "2-HighPoints") > 0 Then
                '   Split the file
                mySplit = Split(arrFixedTested(1), "|")
                '   Write the value
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
            Else
                If InStr(1, rc, "3-MediumPoints") > 0 Then
                    '   Split the file
                    mySplit = Split(arrFixedTested(2), "|")
                    '   Write the value
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
                Else
                    If InStr(1, rc, "4-LowPoints") > 0 Then
                        '   Split the file
                        mySplit = Split(arrFixedTested(3), "|")
                        '   Write the value
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
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
    
    '   See if we need to amend or remove the y axis interval
    iCount = 0
    mySplit = Split(arrFixedTested(0))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrFixedTested(1))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrFixedTested(2))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrFixedTested(3))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    
    '   Open the fixed - tested file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFixedTested.aspx", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    '   Look at the count to decide how we change the y axis interval
    Select Case iCount
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
            strText = Replace(strText, "ReplaceYAxis", "Interval=" & Chr(34) & "1" & Chr(34))
        Case Else
            strText = Replace(strText, "ReplaceYAxis", "")
    End Select
    '   Replace in the file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFixedTested.aspx", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "DefectTimeOpenTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectTimeOpen.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "1-CriticalPoints") > 0 Then
            '   Split the file
            mySplit = Split(arrTimeOpen(0), "|")
                '   Write the value
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
        Else
            If InStr(1, rc, "2-HighPoints") > 0 Then
                '   Split the file
                mySplit = Split(arrTimeOpen(1), "|")
                '   Write the value
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
            Else
                If InStr(1, rc, "3-MediumPoints") > 0 Then
                    '   Split the file
                    mySplit = Split(arrTimeOpen(2), "|")
                    '   Write the value
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
                Else
                    If InStr(1, rc, "4-LowPoints") > 0 Then
                        '   Split the file
                        mySplit = Split(arrTimeOpen(3), "|")
                        '   Write the value
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(0) & Chr(34) & " AxisLabel=" & Chr(34) & "Zero days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(1) & Chr(34) & " AxisLabel=" & Chr(34) & "1 day" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(2) & Chr(34) & " AxisLabel=" & Chr(34) & "2 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(3) & Chr(34) & " AxisLabel=" & Chr(34) & "3 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(4) & Chr(34) & " AxisLabel=" & Chr(34) & "4 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(5) & Chr(34) & " AxisLabel=" & Chr(34) & "5 days" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(6) & Chr(34) & " AxisLabel=" & Chr(34) & "6 day2" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(7) & Chr(34) & " AxisLabel=" & Chr(34) & "1 week" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(8) & Chr(34) & " AxisLabel=" & Chr(34) & "2 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(9) & Chr(34) & " AxisLabel=" & Chr(34) & "3 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(10) & Chr(34) & " AxisLabel=" & Chr(34) & "4 weeks" & Chr(34) & " />"
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & mySplit(11) & Chr(34) & " AxisLabel=" & Chr(34) & "> 4 weeks" & Chr(34) & " />"
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
    
    '   See if we need to amend or remove the y axis interval
    iCount = 0
    mySplit = Split(arrTimeOpen(0))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrTimeOpen(1))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrTimeOpen(2))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    mySplit = Split(arrTimeOpen(3))
    For Each Ele In mySplit
        myNextSplit = Split(Ele, "|")
        For Each NextEle In myNextSplit
            If NextEle <> "0" Then
                iCount = iCount + CInt(NextEle)
            End If
        Next
    Next
    
    '   Open the time open file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectTimeOpen.aspx", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    '   Look at the count to decide how we change the y axis interval
    Select Case iCount
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
            strText = Replace(strText, "ReplaceYAxis", "Interval=" & Chr(34) & "1" & Chr(34))
        Case Else
            strText = Replace(strText, "ReplaceYAxis", "")
    End Select
    '   Replace in the file
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectTimeOpen.aspx", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Now open the table template for new - fixed
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectNewFixedTableTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    '   Replace all the values with our data
    
    '   Critical 1st
    mySplit = Split(arrNewFixed(0), "|")
    strText = Replace(strText, "ZeroCrit", mySplit(0))
    strText = Replace(strText, "1DayCrit", mySplit(1))
    strText = Replace(strText, "2DayCrit", mySplit(2))
    strText = Replace(strText, "3DayCrit", mySplit(3))
    strText = Replace(strText, "4DayCrit", mySplit(4))
    strText = Replace(strText, "5DayCrit", mySplit(5))
    strText = Replace(strText, "6DayCrit", mySplit(6))
    strText = Replace(strText, "1WeekCrit", mySplit(7))
    strText = Replace(strText, "2WeekCrit", mySplit(8))
    strText = Replace(strText, "3WeekCrit", mySplit(9))
    strText = Replace(strText, "4WeekCrit", mySplit(10))
    strText = Replace(strText, "GT4Crit", mySplit(11))
    mySplit = Split(arrNewFixed(1), "|")
    strText = Replace(strText, "ZeroHigh", mySplit(0))
    strText = Replace(strText, "1DayHigh", mySplit(1))
    strText = Replace(strText, "2DayHigh", mySplit(2))
    strText = Replace(strText, "3DayHigh", mySplit(3))
    strText = Replace(strText, "4DayHigh", mySplit(4))
    strText = Replace(strText, "5DayHigh", mySplit(5))
    strText = Replace(strText, "6DayHigh", mySplit(6))
    strText = Replace(strText, "1WeekHigh", mySplit(7))
    strText = Replace(strText, "2WeekHigh", mySplit(8))
    strText = Replace(strText, "3WeekHigh", mySplit(9))
    strText = Replace(strText, "4WeekHigh", mySplit(10))
    strText = Replace(strText, "GT4High", mySplit(11))
    mySplit = Split(arrNewFixed(2), "|")
    strText = Replace(strText, "ZeroMed", mySplit(0))
    strText = Replace(strText, "1DayMed", mySplit(1))
    strText = Replace(strText, "2DayMed", mySplit(2))
    strText = Replace(strText, "3DayMed", mySplit(3))
    strText = Replace(strText, "4DayMed", mySplit(4))
    strText = Replace(strText, "5DayMed", mySplit(5))
    strText = Replace(strText, "6DayMed", mySplit(6))
    strText = Replace(strText, "1WeekMed", mySplit(7))
    strText = Replace(strText, "2WeekMed", mySplit(8))
    strText = Replace(strText, "3WeekMed", mySplit(9))
    strText = Replace(strText, "4WeekMed", mySplit(10))
    strText = Replace(strText, "GT4Med", mySplit(11))
    mySplit = Split(arrNewFixed(3), "|")
    strText = Replace(strText, "ZeroLow", mySplit(0))
    strText = Replace(strText, "1DayLow", mySplit(1))
    strText = Replace(strText, "2DayLow", mySplit(2))
    strText = Replace(strText, "3DayLow", mySplit(3))
    strText = Replace(strText, "4DayLow", mySplit(4))
    strText = Replace(strText, "5DayLow", mySplit(5))
    strText = Replace(strText, "6DayLow", mySplit(6))
    strText = Replace(strText, "1WeekLow", mySplit(7))
    strText = Replace(strText, "2WeekLow", mySplit(8))
    strText = Replace(strText, "3WeekLow", mySplit(9))
    strText = Replace(strText, "4WeekLow", mySplit(10))
    strText = Replace(strText, "GT4Low", mySplit(11))
    
    '   Now open the file for new - fixed
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectNewFixedTable.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Now open the table template for new - fixed
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectFixedTestedTableTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    '   Replace all the values with our data
    
    '   Critical 1st
    mySplit = Split(arrFixedTested(0), "|")
    strText = Replace(strText, "ZeroCrit", mySplit(0))
    strText = Replace(strText, "1DayCrit", mySplit(1))
    strText = Replace(strText, "2DayCrit", mySplit(2))
    strText = Replace(strText, "3DayCrit", mySplit(3))
    strText = Replace(strText, "4DayCrit", mySplit(4))
    strText = Replace(strText, "5DayCrit", mySplit(5))
    strText = Replace(strText, "6DayCrit", mySplit(6))
    strText = Replace(strText, "1WeekCrit", mySplit(7))
    strText = Replace(strText, "2WeekCrit", mySplit(8))
    strText = Replace(strText, "3WeekCrit", mySplit(9))
    strText = Replace(strText, "4WeekCrit", mySplit(10))
    strText = Replace(strText, "GT4Crit", mySplit(11))
    mySplit = Split(arrFixedTested(1), "|")
    strText = Replace(strText, "ZeroHigh", mySplit(0))
    strText = Replace(strText, "1DayHigh", mySplit(1))
    strText = Replace(strText, "2DayHigh", mySplit(2))
    strText = Replace(strText, "3DayHigh", mySplit(3))
    strText = Replace(strText, "4DayHigh", mySplit(4))
    strText = Replace(strText, "5DayHigh", mySplit(5))
    strText = Replace(strText, "6DayHigh", mySplit(6))
    strText = Replace(strText, "1WeekHigh", mySplit(7))
    strText = Replace(strText, "2WeekHigh", mySplit(8))
    strText = Replace(strText, "3WeekHigh", mySplit(9))
    strText = Replace(strText, "4WeekHigh", mySplit(10))
    strText = Replace(strText, "GT4High", mySplit(11))
    mySplit = Split(arrFixedTested(2), "|")
    strText = Replace(strText, "ZeroMed", mySplit(0))
    strText = Replace(strText, "1DayMed", mySplit(1))
    strText = Replace(strText, "2DayMed", mySplit(2))
    strText = Replace(strText, "3DayMed", mySplit(3))
    strText = Replace(strText, "4DayMed", mySplit(4))
    strText = Replace(strText, "5DayMed", mySplit(5))
    strText = Replace(strText, "6DayMed", mySplit(6))
    strText = Replace(strText, "1WeekMed", mySplit(7))
    strText = Replace(strText, "2WeekMed", mySplit(8))
    strText = Replace(strText, "3WeekMed", mySplit(9))
    strText = Replace(strText, "4WeekMed", mySplit(10))
    strText = Replace(strText, "GT4Med", mySplit(11))
    mySplit = Split(arrFixedTested(3), "|")
    strText = Replace(strText, "ZeroLow", mySplit(0))
    strText = Replace(strText, "1DayLow", mySplit(1))
    strText = Replace(strText, "2DayLow", mySplit(2))
    strText = Replace(strText, "3DayLow", mySplit(3))
    strText = Replace(strText, "4DayLow", mySplit(4))
    strText = Replace(strText, "5DayLow", mySplit(5))
    strText = Replace(strText, "6DayLow", mySplit(6))
    strText = Replace(strText, "1WeekLow", mySplit(7))
    strText = Replace(strText, "2WeekLow", mySplit(8))
    strText = Replace(strText, "3WeekLow", mySplit(9))
    strText = Replace(strText, "4WeekLow", mySplit(10))
    strText = Replace(strText, "GT4Low", mySplit(11))
    
    '   Now open the file for new - fixed
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFixedTestedTable.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Now open the table template for new - fixed
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectTimeOpenTableTemplate.txt", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    '   Replace all the values with our data
    
    '   Critical 1st
    mySplit = Split(arrTimeOpen(0), "|")
    strText = Replace(strText, "ZeroCrit", mySplit(0))
    strText = Replace(strText, "1DayCrit", mySplit(1))
    strText = Replace(strText, "2DayCrit", mySplit(2))
    strText = Replace(strText, "3DayCrit", mySplit(3))
    strText = Replace(strText, "4DayCrit", mySplit(4))
    strText = Replace(strText, "5DayCrit", mySplit(5))
    strText = Replace(strText, "6DayCrit", mySplit(6))
    strText = Replace(strText, "1WeekCrit", mySplit(7))
    strText = Replace(strText, "2WeekCrit", mySplit(8))
    strText = Replace(strText, "3WeekCrit", mySplit(9))
    strText = Replace(strText, "4WeekCrit", mySplit(10))
    strText = Replace(strText, "GT4Crit", mySplit(11))
    mySplit = Split(arrTimeOpen(1), "|")
    strText = Replace(strText, "ZeroHigh", mySplit(0))
    strText = Replace(strText, "1DayHigh", mySplit(1))
    strText = Replace(strText, "2DayHigh", mySplit(2))
    strText = Replace(strText, "3DayHigh", mySplit(3))
    strText = Replace(strText, "4DayHigh", mySplit(4))
    strText = Replace(strText, "5DayHigh", mySplit(5))
    strText = Replace(strText, "6DayHigh", mySplit(6))
    strText = Replace(strText, "1WeekHigh", mySplit(7))
    strText = Replace(strText, "2WeekHigh", mySplit(8))
    strText = Replace(strText, "3WeekHigh", mySplit(9))
    strText = Replace(strText, "4WeekHigh", mySplit(10))
    strText = Replace(strText, "GT4High", mySplit(11))
    mySplit = Split(arrTimeOpen(2), "|")
    strText = Replace(strText, "ZeroMed", mySplit(0))
    strText = Replace(strText, "1DayMed", mySplit(1))
    strText = Replace(strText, "2DayMed", mySplit(2))
    strText = Replace(strText, "3DayMed", mySplit(3))
    strText = Replace(strText, "4DayMed", mySplit(4))
    strText = Replace(strText, "5DayMed", mySplit(5))
    strText = Replace(strText, "6DayMed", mySplit(6))
    strText = Replace(strText, "1WeekMed", mySplit(7))
    strText = Replace(strText, "2WeekMed", mySplit(8))
    strText = Replace(strText, "3WeekMed", mySplit(9))
    strText = Replace(strText, "4WeekMed", mySplit(10))
    strText = Replace(strText, "GT4Med", mySplit(11))
    mySplit = Split(arrTimeOpen(3), "|")
    strText = Replace(strText, "ZeroLow", mySplit(0))
    strText = Replace(strText, "1DayLow", mySplit(1))
    strText = Replace(strText, "2DayLow", mySplit(2))
    strText = Replace(strText, "3DayLow", mySplit(3))
    strText = Replace(strText, "4DayLow", mySplit(4))
    strText = Replace(strText, "5DayLow", mySplit(5))
    strText = Replace(strText, "6DayLow", mySplit(6))
    strText = Replace(strText, "1WeekLow", mySplit(7))
    strText = Replace(strText, "2WeekLow", mySplit(8))
    strText = Replace(strText, "3WeekLow", mySplit(9))
    strText = Replace(strText, "4WeekLow", mySplit(10))
    strText = Replace(strText, "GT4Low", mySplit(11))
    
    '   Now open the file for new - fixed
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectTimeOpenTable.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
End Function
Public Function DefectAgeBySeverity(ByVal strType As String)
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

Dim intAge As Integer, intAgeS1 As Integer, intAgeS2 As Integer, intAgeS3 As Integer, intAgeS4 As Integer
Dim intZeroDays As Integer, intZeroDaysS1 As Integer, intZeroDaysS2 As Integer, intZeroDaysS3 As Integer, intZeroDaysS4 As Integer
Dim intOneDay As Integer, intOneDayS1 As Integer, intOneDayS2 As Integer, intOneDayS3 As Integer, intOneDayS4 As Integer
Dim intTwoDays As Integer, intTwoDaysS1 As Integer, intTwoDaysS2 As Integer, intTwoDaysS3 As Integer, intTwoDaysS4 As Integer
Dim intThreeDays As Integer, intThreeDaysS1 As Integer, intThreeDaysS2 As Integer, intThreeDaysS3 As Integer, intThreeDaysS4 As Integer
Dim intFourDays As Integer, intFourDaysS1 As Integer, intFourDaysS2 As Integer, intFourDaysS3 As Integer, intFourDaysS4 As Integer
Dim intFiveDays As Integer, intFiveDaysS1 As Integer, intFiveDaysS2 As Integer, intFiveDaysS3 As Integer, intFiveDaysS4 As Integer
Dim intSixDays As Integer, intSixDaysS1 As Integer, intSixDaysS2 As Integer, intSixDaysS3 As Integer, intSixDaysS4 As Integer
Dim intOneWeek As Integer, intOneWeekS1 As Integer, intOneWeekS2 As Integer, intOneWeekS3 As Integer, intOneWeekS4 As Integer
Dim intTwoWeeks As Integer, intTwoWeeksS1 As Integer, intTwoWeeksS2 As Integer, intTwoWeeksS3 As Integer, intTwoWeeksS4 As Integer
Dim intThreeWeeks As Integer, intThreeWeeksS1 As Integer, intThreeWeeksS2 As Integer, intThreeWeeksS3 As Integer, intThreeWeeksS4 As Integer
Dim intFourWeeks As Integer, intFourWeeksS1 As Integer, intFourWeeksS2 As Integer, intFourWeeksS3 As Integer, intFourWeeksS4 As Integer
Dim intFourWeeksPlus As Integer, intFourWeeksPlusS1 As Integer, intFourWeeksPlusS2 As Integer, intFourWeeksPlusS3 As Integer, intFourWeeksPlusS4 As Integer
Dim strSeverity As String
Dim arrResults(3)
Dim blnDontProcess As Boolean

    '   Set up the main filter
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    tdcBugFilter.Filter("BG_DETECTION_DATE") = "<= " & TodaysDate
    
    '   See if we're doing new - fixed or fixed - tested
    Select Case strType
        Case "New-Fixed"
            tdcBugFilter.Filter("BG_STATUS") = "Fixed Or Tested Or Closed"
        Case "Fixed-Tested"
            tdcBugFilter.Filter("BG_STATUS") = "Tested Or Closed"
        Case Else
            tdcBugFilter.Filter("BG_STATUS") = "New Or Assigned Or Open Or Reopen"
    End Select
    
    '   Create the list
    Set colBugList = tdcBugFilter.NewList
    
    '   Loop round the list getting the values
    For Each objBug In colBugList
        
        '   Find out which type we're doing
        Select Case strType
            Case "New-Fixed"
                '   Get the age
                rc = objBug.Field(strFixedOnDateLabel)
                rd = objBug.Field(strDetectedOnDateLabel)
                If rc = "" Or rd = "" Then
                    Set myFile = fso.OpenTextFile(strOutputPath & "ErrorsLog.txt", ForAppending, True)
                    If rc = "" Then
                        myFile.WriteLine "The 'Fixed On Date' for defect id [" & objBug.ID & "] was empty. Defect age by severity 'New-Fixed' values will be incorrect."
                    End If
                    If rd = "" Then
                        myFile.WriteLine "The 'Detected On Date' for defect id [" & objBug.ID & "] was empty. Defect age by severity 'New-Fixed' values will be incorrect."
                    End If
                    myFile.Close
                    blnDontProcess = True
                Else
                    intAge = objBug.Field(strFixedOnDateLabel) - objBug.Field(strDetectedOnDateLabel)
                    blnDontProcess = False
                End If
            Case "Fixed-Tested"
                '   Get the age
                rc = objBug.Field(strTestedOnDateLabel)
                rd = objBug.Field(strFixedOnDateLabel)
                If rc = "" Or rd = "" Then
                    Set myFile = fso.OpenTextFile(strOutputPath & "ErrorsLog.txt", ForAppending, True)
                    If rc = "" Then
                        myFile.WriteLine "The 'Tested On Date' for defect id [" & objBug.ID & "] was empty. Defect age by severity 'New-Fixed' values will be incorrect."
                    End If
                    If rd = "" Then
                        myFile.WriteLine "The 'Fixed On Date' for defect id [" & objBug.ID & "] was empty. Defect age by severity 'New-Fixed' values will be incorrect."
                    End If
                    myFile.Close
                    blnDontProcess = True
                Else
                    intAge = objBug.Field(strTestedOnDateLabel) - objBug.Field(strFixedOnDateLabel)
                    blnDontProcess = False
                End If
            Case Else
                rd = objBug.Field(strDetectedOnDateLabel)
                If rd = "" Then
                    Set myFile = fso.OpenTextFile(strOutputPath & "ErrorsLog.txt", ForAppending, True)
                    If rd = "" Then
                        myFile.WriteLine "The 'Detected On Date' for defect id [" & objBug.ID & "] was empty. Defect age by severity 'New-Fixed' values will be incorrect."
                    End If
                    myFile.Close
                    blnDontProcess = True
                Else
                    intAge = TodaysDate - objBug.Field(strDetectedOnDateLabel)
                    blnDontProcess = False
                End If
        End Select
        
        '   Only process if all is well
        If blnDontProcess = False Then
        
            '   Get the severity
            strSeverity = objBug.Field("BG_SEVERITY")
    
            '   Count the values
            Select Case intAge
                Case 0
                    Select Case strSeverity
                        Case "1-Critical"
                            intZeroDaysS1 = intZeroDaysS1 + 1
                        Case "2-High"
                            intZeroDaysS2 = intZeroDaysS2 + 1
                        Case "3-Medium"
                            intZeroDaysS3 = intZeroDaysS3 + 1
                        Case "4-Low"
                            intZeroDaysS4 = intZeroDaysS4 + 1
                        Case Else
                            intZeroDays = intZeroDays + 1
                    End Select
                Case 1
                    Select Case strSeverity
                        Case "1-Critical"
                            intOneDayS1 = intOneDayS1 + 1
                        Case "2-High"
                            intOneDayS2 = intOneDayS2 + 1
                        Case "3-Medium"
                            intOneDayS3 = intOneDayS3 + 1
                        Case "4-Low"
                            intOneDayS4 = intOneDayS4 + 1
                        Case Else
                            intOneDay = intOneDay + 1
                    End Select
                Case 2
                    Select Case strSeverity
                        Case "1-Critical"
                            intTwoDaysS1 = intTwoDaysS1 + 1
                        Case "2-High"
                            intTwoDaysS2 = intTwoDaysS2 + 1
                        Case "3-Medium"
                            intTwoDaysS3 = intTwoDaysS3 + 1
                        Case "4-Low"
                            intTwoDaysS4 = intTwoDaysS4 + 1
                        Case Else
                            intTwoDays = intTwoDays + 1
                    End Select
                Case 3
                    Select Case strSeverity
                        Case "1-Critical"
                            intThreeDaysS1 = intThreeDaysS1 + 1
                        Case "2-High"
                            intThreeDaysS2 = intThreeDaysS2 + 1
                        Case "3-Medium"
                            intThreeDaysS3 = intThreeDaysS3 + 1
                        Case "4-Low"
                            intThreeDaysS4 = intThreeDaysS4 + 1
                        Case Else
                            intThreeDays = intThreeDays + 1
                    End Select
                Case 4
                    Select Case strSeverity
                        Case "1-Critical"
                            intFourDaysS1 = intFourDaysS1 + 1
                        Case "2-High"
                            intFourDaysS2 = intFourDaysS2 + 1
                        Case "3-Medium"
                            intFourDaysS3 = intFourDaysS3 + 1
                        Case "4-Low"
                            intFourDaysS4 = intFourDaysS4 + 1
                        Case Else
                            intFourDays = intFourDays + 1
                    End Select
                Case 5
                    Select Case strSeverity
                        Case "1-Critical"
                            intFiveDaysS1 = intFiveDaysS1 + 1
                        Case "2-High"
                            intFiveDaysS2 = intFiveDaysS2 + 1
                        Case "3-Medium"
                            intFiveDaysS3 = intFiveDaysS3 + 1
                        Case "4-Low"
                            intFiveDaysS4 = intFiveDaysS4 + 1
                        Case Else
                            intFiveDays = intFiveDays + 1
                    End Select
                Case 6
                    Select Case strSeverity
                        Case "1-Critical"
                            intSixDaysS1 = intSixDaysS1 + 1
                        Case "2-High"
                            intSixDaysS2 = intSixDaysS2 + 1
                        Case "3-Medium"
                            intSixDaysS3 = intSixDaysS3 + 1
                        Case "4-Low"
                            intSixDaysS4 = intSixDaysS4 + 1
                        Case Else
                            intSixDays = intSixDays + 1
                    End Select
                Case 7
                    Select Case strSeverity
                        Case "1-Critical"
                            intOneWeekS1 = intOneWeekS1 + 1
                        Case "2-High"
                            intOneWeekS2 = intOneWeekS2 + 1
                        Case "3-Medium"
                            intOneWeekS3 = intOneWeekS3 + 1
                        Case "4-Low"
                            intOneWeekS4 = intOneWeekS4 + 1
                        Case Else
                            intOneWeek = intOneWeek + 1
                    End Select
                Case 8, 9, 10, 11, 12, 13, 14
                    Select Case strSeverity
                        Case "1-Critical"
                            intTwoWeeksS1 = intTwoWeeksS1 + 1
                        Case "2-High"
                            intTwoWeeksS2 = intTwoWeeksS2 + 1
                        Case "3-Medium"
                            intTwoWeeksS3 = intTwoWeeksS3 + 1
                        Case "4-Low"
                            intTwoWeeksS4 = intTwoWeeksS4 + 1
                        Case Else
                            intTwoWeeks = intTwoWeeks + 1
                    End Select
                Case 15, 16, 17, 18, 19, 20, 21
                    Select Case strSeverity
                        Case "1-Critical"
                            intThreeWeeksS1 = intThreeWeeksS1 + 1
                        Case "2-High"
                            intThreeWeeksS2 = intThreeWeeksS2 + 1
                        Case "3-Medium"
                            intThreeWeeksS3 = intThreeWeeksS3 + 1
                        Case "4-Low"
                            intThreeWeeksS4 = intThreeWeeksS4 + 1
                        Case Else
                            intThreeWeeks = intThreeWeeks + 1
                    End Select
                Case 22, 23, 24, 25, 26, 27, 28
                    Select Case strSeverity
                        Case "1-Critical"
                            intFourWeeksS1 = intFourWeeksS1 + 1
                        Case "2-High"
                            intFourWeeksS2 = intFourWeeksS2 + 1
                        Case "3-Medium"
                            intFourWeeksS3 = intFourWeeksS3 + 1
                        Case "4-Low"
                            intFourWeeksS4 = intFourWeeksS4 + 1
                        Case Else
                            intFourWeeks = intFourWeeks + 1
                    End Select
                Case Else
                    Select Case strSeverity
                        Case "1-Critical"
                            intFourWeeksPlusS1 = intFourWeeksPlusS1 + 1
                        Case "2-High"
                            intFourWeeksPlusS2 = intFourWeeksPlusS2 + 1
                        Case "3-Medium"
                            intFourWeeksPlusS3 = intFourWeeksPlusS3 + 1
                        Case "4-Low"
                            intFourWeeksPlusS4 = intFourWeeksPlusS4 + 1
                        Case Else
                            intFourWeeksPlus = intFourWeeksPlus + 1
                    End Select
            End Select
        End If
    Next
    
    '   Build the array
    arrResults(0) = intZeroDaysS1 & "|" & intOneDayS1 & "|" & intTwoDaysS1 & "|" & intThreeDaysS1 & "|" & intFourDaysS1 _
    & "|" & intFiveDaysS1 & "|" & intSixDaysS1 & "|" & intOneWeekS1 & "|" & intTwoWeeksS1 & "|" & intThreeWeeksS1 & "|" & intFourWeeksS1 & "|" & intFourWeeksPlusS1
    arrResults(1) = intZeroDaysS2 & "|" & intOneDayS2 & "|" & intTwoDaysS2 & "|" & intThreeDaysS2 & "|" & intFourDaysS2 _
    & "|" & intFiveDaysS2 & "|" & intSixDaysS2 & "|" & intOneWeekS2 & "|" & intTwoWeeksS2 & "|" & intThreeWeeksS2 & "|" & intFourWeeksS2 & "|" & intFourWeeksPlusS2
    arrResults(2) = intZeroDaysS3 & "|" & intOneDayS3 & "|" & intTwoDaysS3 & "|" & intThreeDaysS3 & "|" & intFourDaysS3 _
    & "|" & intFiveDaysS3 & "|" & intSixDaysS3 & "|" & intOneWeekS3 & "|" & intTwoWeeksS3 & "|" & intThreeWeeksS3 & "|" & intFourWeeksS3 & "|" & intFourWeeksPlusS3
    arrResults(3) = intZeroDaysS4 & "|" & intOneDayS4 & "|" & intTwoDaysS4 & "|" & intThreeDaysS4 & "|" & intFourDaysS4 _
    & "|" & intFiveDaysS4 & "|" & intSixDaysS4 & "|" & intOneWeekS4 & "|" & intTwoWeeksS4 & "|" & intThreeWeeksS4 & "|" & intFourWeeksS4 & "|" & intFourWeeksPlusS4
    
    DefectAgeBySeverity = arrResults
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function GetDefectStatusBySeverity()

    arrStatusBySeverity = DefectStatusBySeverity()
    iRow = 5
    '   Write out the 1st section
    For i = LBound(arrStatusBySeverity) To 3
        iCol = 2
        mySplit = Split(arrStatusBySeverity(i), "|")
        For j = LBound(mySplit) To UBound(mySplit)
            objWrkBk.Worksheets(wrkShtDefectProgress).Cells(iRow, iCol).Value = mySplit(j)
            iCol = iCol + 1
        Next
        iRow = iRow + 1
    Next
    iRow = 11
    '   Write out the remaining section
    For i = 4 To UBound(arrStatusBySeverity)
        iCol = 2
        mySplit = Split(arrStatusBySeverity(i), "|")
        For j = LBound(mySplit) To UBound(mySplit)
            objWrkBk.Worksheets(wrkShtDefectProgress).Cells(iRow, iCol).Value = mySplit(j)
            iCol = iCol + 1
        Next
        iRow = iRow + 1
    Next
    
End Function
Public Function GetDefectStatusByPriority()

    arrStatusByPriority = DefectStatusByPriority()
    iRow = 5
    '   Write out the 1st Section
    For i = LBound(arrStatusByPriority) To 3
        iCol = 9
        mySplit = Split(arrStatusByPriority(i), "|")
        For j = LBound(mySplit) To UBound(mySplit)
            objWrkBk.Worksheets(wrkShtDefectProgress).Cells(iRow, iCol).Value = mySplit(j)
            iCol = iCol + 1
        Next
        iRow = iRow + 1
    Next
    iRow = 11
    '   Now the last section
    For i = 4 To UBound(arrStatusByPriority)
        iCol = 9
        mySplit = Split(arrStatusByPriority(i), "|")
        For j = LBound(mySplit) To UBound(mySplit)
            objWrkBk.Worksheets(wrkShtDefectProgress).Cells(iRow, iCol).Value = mySplit(j)
            iCol = iCol + 1
        Next
        iRow = iRow + 1
    Next
    
End Function
Public Function DefectStatusBySeverity()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    Dim strStatus As String, strSeverity As String
    Dim arrReturn(11)
    Dim intNew As Integer, intAssigned As Integer, intOpen As Integer, intFixed As Integer, intTested As Integer, intReopen As Integer, intClosed As Integer, intDuplicate As Integer, intRejected As Integer, intOnHold As Integer, intReady As Integer, intFailed As Integer
    Dim intNewS1 As Integer, intAssignedS1 As Integer, intOpenS1 As Integer, intFixedS1 As Integer, intTestedS1 As Integer, intReopenS1 As Integer, intClosedS1 As Integer, intDuplicateS1 As Integer, intRejectedS1 As Integer, intOnHoldS1 As Integer, intReadyS1 As Integer, intFailedS1 As Integer
    Dim intNewS2 As Integer, intAssignedS2 As Integer, intOpenS2 As Integer, intFixedS2 As Integer, intTestedS2 As Integer, intReopenS2 As Integer, intClosedS2 As Integer, intDuplicateS2 As Integer, intRejectedS2 As Integer, intOnHoldS2 As Integer, intReadyS2 As Integer, intFailedS2 As Integer
    Dim intNewS3 As Integer, intAssignedS3 As Integer, intOpenS3 As Integer, intFixedS3 As Integer, intTestedS3 As Integer, intReopenS3 As Integer, intClosedS3 As Integer, intDuplicateS3 As Integer, intRejectedS3 As Integer, intOnHoldS3 As Integer, intReadyS3 As Integer, intFailedS3 As Integer
    Dim intNewS4 As Integer, intAssignedS4 As Integer, intOpenS4 As Integer, intFixedS4 As Integer, intTestedS4 As Integer, intReopenS4 As Integer, intClosedS4 As Integer, intDuplicateS4 As Integer, intRejectedS4 As Integer, intOnHoldS4 As Integer, intReadyS4 As Integer, intFailedS4 As Integer
    
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    tdcBugFilter.Filter("BG_DETECTION_DATE") = "<= " & TodaysDate
   
    Set colBugList = tdcBugFilter.NewList
    
    For Each objBug In colBugList
        strStatus = objBug.Field("BG_STATUS")
        strSeverity = objBug.Field("BG_SEVERITY")
        Select Case strStatus
            Case "New"
                Select Case strSeverity
                    Case "1-Critical"
                        intNewS1 = intNewS1 + 1
                    Case "2-High"
                        intNewS2 = intNewS2 + 1
                    Case "3-Medium"
                        intNewS3 = intNewS3 + 1
                    Case "4-Low"
                        intNewS4 = intNewS4 + 1
                    Case Else
                        intNew = intNew + 1
                End Select
            Case "Assigned"
                Select Case strSeverity
                    Case "1-Critical"
                        intAssignedS1 = intAssignedS1 + 1
                    Case "2-High"
                        intAssignedS2 = intAssignedS2 + 1
                    Case "3-Medium"
                        intAssignedS3 = intAssignedS3 + 1
                    Case "4-Low"
                        intAssignedS4 = intAssignedS4 + 1
                    Case Else
                        intAssigned = intAssigned + 1
                End Select
            Case "Open"
                Select Case strSeverity
                    Case "1-Critical"
                        intOpenS1 = intOpenS1 + 1
                    Case "2-High"
                        intOpenS2 = intOpenS2 + 1
                    Case "3-Medium"
                        intOpenS3 = intOpenS3 + 1
                    Case "4-Low"
                        intOpenS4 = intOpenS4 + 1
                    Case Else
                        intOpen = intOpen + 1
                End Select
            Case "Fixed"
                Select Case strSeverity
                    Case "1-Critical"
                        intFixedS1 = intFixedS1 + 1
                    Case "2-High"
                        intFixedS2 = intFixedS2 + 1
                    Case "3-Medium"
                        intFixedS3 = intFixedS3 + 1
                    Case "4-Low"
                        intFixedS4 = intFixedS4 + 1
                    Case Else
                        intFixed = intFixed + 1
                End Select
            Case "Ready For Testing", "Ready for Testing"
                Select Case strSeverity
                    Case "1-Critical"
                        intReadyS1 = intReadyS1 + 1
                    Case "2-High"
                        intReadyS2 = intReadyS2 + 1
                    Case "3-Medium"
                        intReadyS3 = intReadyS3 + 1
                    Case "4-Low"
                        intReadyS4 = intReadyS4 + 1
                    Case Else
                        intReady = intReady + 1
                End Select
            Case "Failed Testing"
                Select Case strSeverity
                    Case "1-Critical"
                        intFailedS1 = intFailedS1 + 1
                    Case "2-High"
                        intFailedS2 = intFailedS2 + 1
                    Case "3-Medium"
                        intFailedS3 = intFailedS3 + 1
                    Case "4-Low"
                        intFailedS4 = intFailedS4 + 1
                    Case Else
                        intFailed = intFailed + 1
                End Select
            Case "Tested"
                Select Case strSeverity
                    Case "1-Critical"
                        intTestedS1 = intTestedS1 + 1
                    Case "2-High"
                        intTestedS2 = intTestedS2 + 1
                    Case "3-Medium"
                        intTestedS3 = intTestedS3 + 1
                    Case "4-Low"
                        intTestedS4 = intTestedS4 + 1
                    Case Else
                        intTested = intTested + 1
                End Select
            Case "Reopen"
                Select Case strSeverity
                    Case "1-Critical"
                        intReopenS1 = intReopenS1 + 1
                    Case "2-High"
                        intReopenS2 = intReopenS2 + 1
                    Case "3-Medium"
                        intReopenS3 = intReopenS3 + 1
                    Case "4-Low"
                        intReopenS4 = intReopenS4 + 1
                    Case Else
                        intReopen = intReopen + 1
                End Select
            Case "Closed"
                Select Case strSeverity
                    Case "1-Critical"
                        intClosedS1 = intClosedS1 + 1
                    Case "2-High"
                        intClosedS2 = intClosedS2 + 1
                    Case "3-Medium"
                        intClosedS3 = intClosedS3 + 1
                    Case "4-Low"
                        intClosedS4 = intClosedS4 + 1
                    Case Else
                        intClosed = intClosed + 1
                End Select
            Case "Duplicate"
                Select Case strSeverity
                    Case "1-Critical"
                        intDuplicateS1 = intDuplicateS1 + 1
                    Case "2-High"
                        intDuplicateS2 = intDuplicateS2 + 1
                    Case "3-Medium"
                        intDuplicateS3 = intDuplicateS3 + 1
                    Case "4-Low"
                        intDuplicateS4 = intDuplicateS4 + 1
                    Case Else
                        intDuplicate = intDuplicate + 1
                End Select
            Case "Rejected"
                Select Case strSeverity
                    Case "1-Critical"
                        intRejectedS1 = intRejectedS1 + 1
                    Case "2-High"
                        intRejectedS2 = intRejectedS2 + 1
                    Case "3-Medium"
                        intRejectedS3 = intRejectedS3 + 1
                    Case "4-Low"
                        intRejectedS4 = intRejectedS4 + 1
                    Case Else
                        intRejected = intRejected + 1
                End Select
            Case "On Hold"
                Select Case strSeverity
                    Case "1-Critical"
                        intOnHoldS1 = intOnHoldS1 + 1
                    Case "2-High"
                        intOnHoldS2 = intOnHoldS2 + 1
                    Case "3-Medium"
                        intOnHoldS3 = intOnHoldS3 + 1
                    Case "4-Low"
                        intOnHoldS4 = intOnHoldS4 + 1
                    Case Else
                        intOnHold = intOnHold + 1
                End Select
        End Select
    Next
    
    '   Add to the array
    arrReturn(0) = "New" & "|" & intNewS1 & "|" & intNewS2 & "|" & intNewS3 & "|" & intNewS4
    arrReturn(1) = "Assigned" & "|" & intAssignedS1 & "|" & intAssignedS2 & "|" & intAssignedS3 & "|" & intAssignedS4
    arrReturn(2) = "Open" & "|" & intOpenS1 & "|" & intOpenS2 & "|" & intOpenS3 & "|" & intOpenS4
    arrReturn(3) = "Reopen" & "|" & intReopenS1 & "|" & intReopenS2 & "|" & intReopenS3 & "|" & intReopenS4
    arrReturn(4) = "Failed Testing" & "|" & intFailedS1 & "|" & intFailedS2 & "|" & intFailedS3 & "|" & intFailedS4
    arrReturn(5) = "Fixed" & "|" & intFixedS1 & "|" & intFixedS2 & "|" & intFixedS3 & "|" & intFixedS4
    arrReturn(6) = "Ready For Testing" & "|" & intReadyS1 & "|" & intReadyS2 & "|" & intReadyS3 & "|" & intReadyS4
    arrReturn(7) = "Tested" & "|" & intTestedS1 & "|" & intTestedS2 & "|" & intTestedS3 & "|" & intTestedS4
    arrReturn(8) = "Duplicate" & "|" & intDuplicateS1 & "|" & intDuplicateS2 & "|" & intDuplicateS3 & "|" & intDuplicateS4
    arrReturn(9) = "Rejected" & "|" & intRejectedS1 & "|" & intRejectedS2 & "|" & intRejectedS3 & "|" & intRejectedS4
    arrReturn(10) = "On Hold" & "|" & intOnHoldS1 & "|" & intOnHoldS2 & "|" & intOnHoldS3 & "|" & intOnHoldS4
    arrReturn(11) = "Closed" & "|" & intClosedS1 & "|" & intClosedS2 & "|" & intClosedS3 & "|" & intClosedS4
    
    '    Return the array
    DefectStatusBySeverity = arrReturn
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function DefectStatusByPriority()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    Dim strStatus As String, strSeverity As String
    Dim arrReturn(11)
    Dim intNew As Integer, intAssigned As Integer, intOpen As Integer, intFixed As Integer, intTested As Integer, intReopen As Integer, intClosed As Integer, intDuplicate As Integer, intRejected As Integer, intOnHold As Integer, intReady As Integer, intFailed As Integer
    Dim intNewS1 As Integer, intAssignedS1 As Integer, intOpenS1 As Integer, intFixedS1 As Integer, intTestedS1 As Integer, intReopenS1 As Integer, intClosedS1 As Integer, intDuplicateS1 As Integer, intRejectedS1 As Integer, intOnHoldS1 As Integer, intReadyS1 As Integer, intFailedS1 As Integer
    Dim intNewS2 As Integer, intAssignedS2 As Integer, intOpenS2 As Integer, intFixedS2 As Integer, intTestedS2 As Integer, intReopenS2 As Integer, intClosedS2 As Integer, intDuplicateS2 As Integer, intRejectedS2 As Integer, intOnHoldS2 As Integer, intReadyS2 As Integer, intFailedS2 As Integer
    Dim intNewS3 As Integer, intAssignedS3 As Integer, intOpenS3 As Integer, intFixedS3 As Integer, intTestedS3 As Integer, intReopenS3 As Integer, intClosedS3 As Integer, intDuplicateS3 As Integer, intRejectedS3 As Integer, intOnHoldS3 As Integer, intReadyS3 As Integer, intFailedS3 As Integer
    Dim intNewS4 As Integer, intAssignedS4 As Integer, intOpenS4 As Integer, intFixedS4 As Integer, intTestedS4 As Integer, intReopenS4 As Integer, intClosedS4 As Integer, intDuplicateS4 As Integer, intRejectedS4 As Integer, intOnHoldS4 As Integer, intReadyS4 As Integer, intFailedS4 As Integer
    
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    tdcBugFilter.Filter("BG_DETECTION_DATE") = "<= " & TodaysDate
   
    Set colBugList = tdcBugFilter.NewList
    
    For Each objBug In colBugList
        strStatus = objBug.Field("BG_STATUS")
        strPriority = objBug.Field("BG_PRIORITY")
        Select Case strStatus
            Case "New"
                Select Case strPriority
                    Case "1-Urgent"
                        intNewS1 = intNewS1 + 1
                    Case "2-High"
                        intNewS2 = intNewS2 + 1
                    Case "3-Medium"
                        intNewS3 = intNewS3 + 1
                    Case "4-Low"
                        intNewS4 = intNewS4 + 1
                    Case Else
                        intNew = intNew + 1
                End Select
            Case "Assigned"
                Select Case strPriority
                    Case "1-Urgent"
                        intAssignedS1 = intAssignedS1 + 1
                    Case "2-High"
                        intAssignedS2 = intAssignedS2 + 1
                    Case "3-Medium"
                        intAssignedS3 = intAssignedS3 + 1
                    Case "4-Low"
                        intAssignedS4 = intAssignedS4 + 1
                    Case Else
                        intAssigned = intAssigned + 1
                End Select
            Case "Open"
                Select Case strPriority
                    Case "1-Urgent"
                        intOpenS1 = intOpenS1 + 1
                    Case "2-High"
                        intOpenS2 = intOpenS2 + 1
                    Case "3-Medium"
                        intOpenS3 = intOpenS3 + 1
                    Case "4-Low"
                        intOpenS4 = intOpenS4 + 1
                    Case Else
                        intOpen = intOpen + 1
                End Select
            Case "Fixed"
                Select Case strPriority
                    Case "1-Urgent"
                        intFixedS1 = intFixedS1 + 1
                    Case "2-High"
                        intFixedS2 = intFixedS2 + 1
                    Case "3-Medium"
                        intFixedS3 = intFixedS3 + 1
                    Case "4-Low"
                        intFixedS4 = intFixedS4 + 1
                    Case Else
                        intFixed = intFixed + 1
                End Select
            Case "Ready For Testing", "Ready for Testing"
                Select Case strPriority
                    Case "1-Urgent"
                        intReadyS1 = intReadyS1 + 1
                    Case "2-High"
                        intReadyS2 = intReadyS2 + 1
                    Case "3-Medium"
                        intReadyS3 = intReadyS3 + 1
                    Case "4-Low"
                        intReadyS4 = intReadyS4 + 1
                    Case Else
                        intReady = intReady + 1
                End Select
            Case "Failed Testing"
                Select Case strPriority
                    Case "1-Urgent"
                        intFailedS1 = intFailedS1 + 1
                    Case "2-High"
                        intFailedS2 = intFailedS2 + 1
                    Case "3-Medium"
                        intFailedS3 = intFailedS3 + 1
                    Case "4-Low"
                        intFailedS4 = intFailedS4 + 1
                    Case Else
                        intFailed = intFailed + 1
                End Select
            Case "Tested"
                Select Case strPriority
                    Case "1-Urgent"
                        intTestedS1 = intTestedS1 + 1
                    Case "2-High"
                        intTestedS2 = intTestedS2 + 1
                    Case "3-Medium"
                        intTestedS3 = intTestedS3 + 1
                    Case "4-Low"
                        intTestedS4 = intTestedS4 + 1
                    Case Else
                        intTested = intTested + 1
                End Select
            Case "Reopen"
                Select Case strPriority
                    Case "1-Urgent"
                        intReopenS1 = intReopenS1 + 1
                    Case "2-High"
                        intReopenS2 = intReopenS2 + 1
                    Case "3-Medium"
                        intReopenS3 = intReopenS3 + 1
                    Case "4-Low"
                        intReopenS4 = intReopenS4 + 1
                    Case Else
                        intReopen = intReopen + 1
                End Select
            Case "Closed"
                Select Case strPriority
                    Case "1-Urgent"
                        intClosedS1 = intClosedS1 + 1
                    Case "2-High"
                        intClosedS2 = intClosedS2 + 1
                    Case "3-Medium"
                        intClosedS3 = intClosedS3 + 1
                    Case "4-Low"
                        intClosedS4 = intClosedS4 + 1
                    Case Else
                        intClosed = intClosed + 1
                End Select
            Case "Duplicate"
                Select Case strPriority
                    Case "1-Urgent"
                        intDuplicateS1 = intDuplicateS1 + 1
                    Case "2-High"
                        intDuplicateS2 = intDuplicateS2 + 1
                    Case "3-Medium"
                        intDuplicateS3 = intDuplicateS3 + 1
                    Case "4-Low"
                        intDuplicateS4 = intDuplicateS4 + 1
                    Case Else
                        intDuplicate = intDuplicate + 1
                End Select
            Case "Rejected"
                Select Case strPriority
                    Case "1-Urgent"
                        intRejectedS1 = intRejectedS1 + 1
                    Case "2-High"
                        intRejectedS2 = intRejectedS2 + 1
                    Case "3-Medium"
                        intRejectedS3 = intRejectedS3 + 1
                    Case "4-Low"
                        intRejectedS4 = intRejectedS4 + 1
                    Case Else
                        intRejected = intRejected + 1
                End Select
            Case "On Hold"
                Select Case strPriority
                    Case "1-Urgent"
                        intOnHoldS1 = intOnHoldS1 + 1
                    Case "2-High"
                        intOnHoldS2 = intOnHoldS2 + 1
                    Case "3-Medium"
                        intOnHoldS3 = intOnHoldS3 + 1
                    Case "4-Low"
                        intOnHoldS4 = intOnHoldS4 + 1
                    Case Else
                        intOnHold = intOnHold + 1
                End Select
        End Select
    Next
    
    '   Add to the array
    arrReturn(0) = "New" & "|" & intNewS1 & "|" & intNewS2 & "|" & intNewS3 & "|" & intNewS4
    arrReturn(1) = "Assigned" & "|" & intAssignedS1 & "|" & intAssignedS2 & "|" & intAssignedS3 & "|" & intAssignedS4
    arrReturn(2) = "Open" & "|" & intOpenS1 & "|" & intOpenS2 & "|" & intOpenS3 & "|" & intOpenS4
    arrReturn(3) = "Reopen" & "|" & intReopenS1 & "|" & intReopenS2 & "|" & intReopenS3 & "|" & intReopenS4
    arrReturn(4) = "Failed Testing" & "|" & intFailedS1 & "|" & intFailedS2 & "|" & intFailedS3 & "|" & intFailedS4
    arrReturn(5) = "Fixed" & "|" & intFixedS1 & "|" & intFixedS2 & "|" & intFixedS3 & "|" & intFixedS4
    arrReturn(6) = "Ready For Testing" & "|" & intReadyS1 & "|" & intReadyS2 & "|" & intReadyS3 & "|" & intReadyS4
    arrReturn(7) = "Tested" & "|" & intTestedS1 & "|" & intTestedS2 & "|" & intTestedS3 & "|" & intTestedS4
    arrReturn(8) = "Duplicate" & "|" & intDuplicateS1 & "|" & intDuplicateS2 & "|" & intDuplicateS3 & "|" & intDuplicateS4
    arrReturn(9) = "Rejected" & "|" & intRejectedS1 & "|" & intRejectedS2 & "|" & intRejectedS3 & "|" & intRejectedS4
    arrReturn(10) = "On Hold" & "|" & intOnHoldS1 & "|" & intOnHoldS2 & "|" & intOnHoldS3 & "|" & intOnHoldS4
    arrReturn(11) = "Closed" & "|" & intClosedS1 & "|" & intClosedS2 & "|" & intClosedS3 & "|" & intClosedS4
    
    '    Return the array
    DefectStatusByPriority = arrReturn
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
End Function
Public Function GetOpenDefectDetails()
Dim strStatus As String, strPriority As String, strSeverity As String, strSummary As String
Dim strDetectedOnDate As String, strAssignedTo As String, strDefectId As String
Dim iRow As Integer

    '   Set the row start point
    iRow = 5

    '   Set up the project and test phase filter
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    tdcBugFilter.Filter("BG_STATUS") = "New Or Assigned Or Open Or Reopen"
    tdcBugFilter.Filter("BG_DETECTION_DATE") = "<= " & TodaysDate
   
    Set colBugList = tdcBugFilter.NewList
    
    '   Get the details by severity
    '   1-Critical
    For Each objBug In colBugList
        strSeverity = objBug.Field("BG_SEVERITY")
        
        If strSeverity = "1-Critical" Then
            '   Get the defect details
            strDefectId = objBug.ID
            strSummary = objBug.Summary
            strDetectedOnDate = objBug.Field(strDetectedOnDateLabel)
            strStatus = objBug.Status
            strAssignedTo = objBug.AssignedTo
            
            strPriority = objBug.Priority
            
            '   Write out the defect details
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 1).Value = strDefectId
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 2).Value = strSummary
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 3).Value = strDetectedOnDate
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 4).Value = strStatus
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 5).Value = strAssignedTo
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 6).Value = strSeverity
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 7).Value = strPriority
            
            '   Up the row
            iRow = iRow + 1
        End If
    Next
    '   2-High
    For Each objBug In colBugList
        strSeverity = objBug.Field("BG_SEVERITY")
        
        If strSeverity = "2-High" Then
            '   Get the defect details
            strDefectId = objBug.ID
            strSummary = objBug.Summary
            strDetectedOnDate = objBug.Field(strDetectedOnDateLabel)
            strStatus = objBug.Status
            strAssignedTo = objBug.AssignedTo
            
            strPriority = objBug.Priority
            
            '   Write out the defect details
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 1).Value = strDefectId
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 2).Value = strSummary
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 3).Value = strDetectedOnDate
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 4).Value = strStatus
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 5).Value = strAssignedTo
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 6).Value = strSeverity
            objWrkBk.Worksheets(wrkShtOpenDefects).Cells(iRow, 7).Value = strPriority
            
            '   Up the row
            iRow = iRow + 1
        End If
    Next
    
End Function
Public Function FindFirstBugDate()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
Dim hst As TDAPIOLELib.History
Dim hstRec As TDAPIOLELib.HistoryRecord
Dim hstList As TDAPIOLELib.list
Dim iCount
Dim myArr()
Dim myTemp As Date
Dim myDate As Date

    '   Set up the main filter
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    
    Set colBugList = tdcBugFilter.NewList
    
    '   Make sure we've got some bugs for this project and test phase
    If colBugList.Count = 0 Then
        FindFirstBugDate = "00:00:00"
        Exit Function
    End If
    
    For Each objBug In colBugList
        myDate = objBug.Field("BG_DETECTION_DATE")
        If myTemp = "00:00:00" Then
            myTemp = myDate
        Else
            If myDate < myTemp Then
                myTemp = myDate
            End If
        End If
    Next
    
    '   Return date
    FindFirstBugDate = myTemp
      
End Function
Public Function GetDefectsByFunctionalArea()
Dim iCount As Integer
Dim iTotCount As Integer
iCount = 0
iTotCount = 0

    ' Count the critical defects by test phase and add them to the data sheet.
    arrTestPhaseKeys = DefectsByFunctionalArea()
    
    '   Build the chart stuff into the template
    
    '  Open the template file
    Set mySource = fso.OpenTextFile(strTemplatePath & "DefectFunctionalAreaTemplate.txt", ForReading)
    Set myDest = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFunctionalArea.aspx", ForWriting, True)
    Do While mySource.AtEndOfStream <> True
        rc = mySource.ReadLine
        If InStr(1, rc, "1-CriticalPoints") > 0 Then
            '   Loop round our array
            For i = 0 To UBound(arrTestPhaseKeys)
                '   Split the file
                mySplit = Split(arrTestPhaseKeys(i), "|")
                If mySplit(0) = "1-Critical" Then
                    '   Write the values
                    myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objBugFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                End If
            Next
        Else
            If InStr(1, rc, "2-HighPoints") > 0 Then
                '   Loop round our array
                For i = 0 To UBound(arrTestPhaseKeys)
                    '   Split the file
                    mySplit = Split(arrTestPhaseKeys(i), "|")
                    If mySplit(0) = "2-High" Then
                        '   Write the values
                        myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objBugFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                    End If
                Next
            Else
                If InStr(1, rc, "3-MediumPoints") > 0 Then
                    '   Loop round our array
                    For i = 0 To UBound(arrTestPhaseKeys)
                        '   Split the file
                        mySplit = Split(arrTestPhaseKeys(i), "|")
                        If mySplit(0) = "3-Medium" Then
                            '   Write the values
                            myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objBugFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
                        End If
                    Next
                Else
                    If InStr(1, rc, "4-LowPoints") > 0 Then
                        '   Loop round our array
                        For i = 0 To UBound(arrTestPhaseKeys)
                            '   Split the file
                            mySplit = Split(arrTestPhaseKeys(i), "|")
                            If mySplit(0) = "4-Low" Then
                                '   Write the values
                                myDest.WriteLine "<DCWC:DataPoint YValues=" & Chr(34) & objBugFunctionalAreaDictionary(arrTestPhaseKeys(i)) & Chr(34) & " AxisLabel=" & Chr(34) & mySplit(1) & Chr(34) & " />"
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
        If mySplit(0) = "1-Critical" Then
            If myTemp = "" Or myTemp <> mySplit(1) Then
                myTemp = mySplit(1)
                iCount = iCount + 1
            End If
        End If
    Next
    myTemp = ""
    For Each Ele In arrTestPhaseKeys
        iTotCount = iTotCount + objBugFunctionalAreaDictionary(Ele)
    Next
    
    '   Open the file and change the width value
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFunctionalArea.aspx", ForReading)
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
    
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFunctionalArea.aspx", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
    '   Copy the hyperlink table file to the folder
    fso.CopyFile strTemplatePath & "DefectFunctionalAreaTableTemplate.txt", strFolderPath & strPathandFileName & "-DefectFunctionalAreaTable.asp"
    '   Now do the hyperlink table
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectFunctionalAreaTable.asp", ForAppending)
    myTemp = ""
    For i = 0 To UBound(arrTestPhaseKeys)
        '   Split the file
        mySplit = Split(arrTestPhaseKeys(i), "|")
        If mySplit(0) = "1-Critical" Then
            If myTemp = "" Or myTemp <> mySplit(1) Then
                myTemp = mySplit(1)
                myFile.WriteLine "<tr><td>" & mySplit(1) & "</td><td>" & objBugFunctionalAreaDictionary(arrTestPhaseKeys(i)) & "</td><td>" & objBugFunctionalAreaDictionary("2-High|" & mySplit(1)) & "</td><td>" & objBugFunctionalAreaDictionary("3-Medium|" & mySplit(1)) & "</td><td>" & objBugFunctionalAreaDictionary("4-Low|" & mySplit(1)) & "</td></tr>"
            End If
        Else
            Exit For
        End If
    Next
    '   Finish off html
    myFile.WriteLine "</table></table></body></html>"
    myFile.Close
    
End Function
Public Function DefectsByFunctionalArea()
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
Dim strSeverity As String
Dim arrSeverity(3)
Dim blnOK As Boolean
    
    arrSeverity(0) = "1-Critical"
    arrSeverity(1) = "2-High"
    arrSeverity(2) = "3-Medium"
    arrSeverity(3) = "4-Low"

    ' Create the dictionary object
    Set objBugFunctionalAreaDictionary = New Dictionary
    
    '   Create the filter
    Set tdcBugFactory = tdc.BugFactory
    Set tdcBugFilter = tdcBugFactory.Filter
    tdcBugFilter.Filter("BG_PROJECT") = Chr(39) & strProjectName & Chr(39)
    tdcBugFilter.Filter(strTestPhaseBugLabel) = Chr(39) & strTestPhase & Chr(39)
    '   See if we're filtering on Sub-Project
    If strSubProjectName <> "N/A" Then
        tdcBugFilter.Filter(strSubProjectBugLabel) = Chr(34) & strSubProjectName & Chr(34)
    End If
    
    '   Get the list
    Set colBugList = tdcBugFilter.NewList
    
    '   Go through each functional area in the list and see if we have any stats for it
    For Each objBug In colBugList
        strThisFunctionalArea = objBug.Field(strFunctionalAreaBugLabel)
        If strThisFunctionalArea = "" Then
            strThisFunctionalArea = "Not Assigned"
        End If
        
        If objBugFunctionalAreaDictionary.Exists(arrSeverity(0) & "|" & strThisFunctionalArea) = True Then
            blnOK = True
        Else
            If objBugFunctionalAreaDictionary.Exists(arrSeverity(1) & "|" & strThisFunctionalArea) = True Then
                blnOK = True
            Else
                If objBugFunctionalAreaDictionary.Exists(arrSeverity(2) & "|" & strThisFunctionalArea) = True Then
                    blnOK = True
                Else
                    If objBugFunctionalAreaDictionary.Exists(arrSeverity(3) & "|" & strThisFunctionalArea) = True Then
                        blnOK = True
                    Else
                        blnOK = False
                    End If
                End If
            End If
        End If
        '   See if we've not got the item in our list
        If blnOK = False Then
            '   Add the new item to the dictionary with its severities
            objBugFunctionalAreaDictionary.Add arrSeverity(0) & "|" & strThisFunctionalArea, 0
            objBugFunctionalAreaDictionary.Add arrSeverity(1) & "|" & strThisFunctionalArea, 0
            objBugFunctionalAreaDictionary.Add arrSeverity(2) & "|" & strThisFunctionalArea, 0
            objBugFunctionalAreaDictionary.Add arrSeverity(3) & "|" & strThisFunctionalArea, 0
        End If
        strSeverity = objBug.Field("BG_SEVERITY")
        intCount = CInt(objBugFunctionalAreaDictionary.Item(strSeverity & "|" & strThisFunctionalArea))
        objBugFunctionalAreaDictionary.Item(strSeverity & "|" & strThisFunctionalArea) = intCount + 1
        blnOK = False
    Next
    
    '   Sort the array
    SortDictionary objBugFunctionalAreaDictionary, True

    DefectsByFunctionalArea = objBugFunctionalAreaDictionary.Keys
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
End Function

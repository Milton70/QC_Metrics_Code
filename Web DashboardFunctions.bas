Attribute VB_Name = "DashboardFunctions"
Public Function GetListOfSubProjects()

On Error GoTo ErrorHandler

    Dim arListValues() As String
    Dim arLabels As Variant
    Dim iCount As Integer
    Dim strListName As String
    Dim strFieldName As String
    
    strFieldName = GetFieldName("Sub Project", "CYCLE")
    strListName = customLists(strFieldName, "CYCLE")
    arListValues = GetComboValues(strListName)
    
    '   Return the array list
    GetListOfSubProjects = arListValues
    
    Exit Function
    
ErrorHandler:

    Call ErrorHandler(Err)
    Resume Next
    
End Function
Public Function GetListOfTestPhases()

On Error GoTo ErrorHandler

    Dim arListValues() As String
    Dim arLabels As Variant
    Dim iCount As Integer
    Dim strListName As String
    Dim strFieldName As String
    
    strFieldName = GetFieldName("Test Phase", "BUG")
    strListName = customLists(strFieldName, "BUG")
    arListValues = GetComboValues(strListName)
    
    '   Return the array
    GetListOfTestPhases = arListValues
    
    Exit Function
    
ErrorHandler:

    Call ErrorHandler(Err)
    Resume Next
    
End Function
Public Function AreWeUsingSubProjects() As Boolean

    '   Get all test set within the project
    Set tdcTestSetFactory = tdc.TestSetFactory
    
     ' Get the test set tree manager object.
    Set objTestSetTreeManager = tdc.TestSetTreeManager
    
    ' Find all the test sets under the root.
    Set objTestSetFolder = objTestSetTreeManager.Root
    Set colTestSets = objTestSetFolder.FindTestSets("")
    
    '   Go through each test set
    For Each objTestSet In colTestSets
        '   Set up the filter by project
        If objTestSet.Field(strProjectCycleLabel) = strProjectName Then
            '   Get the Sub-project
            strSubProjectName = objTestSet.Field(strSubProjectCycleLabel)
            '   See if we've got a value or just empty
            If strSubProjectName <> "" Then
                AreWeUsingSubProjects = True
                Exit Function
            End If
        End If
    Next
   
    '   If we get here then it's false and we're not doing sub-projects
    AreWeUsingSubProjects = False

End Function
Public Function GetBaselineDate()
Dim strReturnText As String

    '   Get all the dates
    myBaselineDate = FindLastDate("Planned-Baseline")
    myReplannedDate = FindLastDate("Replanned")
    myActualDate = FindLastDate("Actual")
    '   See if we've got any dates
    If myBaselineDate = "00:00:00" Then
        If myReplannedDate = "00:00:00" Then
            If myActualDate = "00:00:00" Then
                '   We've go no dates so we ignore this test phase
                GetBaselineDate = "False"
                Exit Function
            End If
        End If
    End If
    
    '   See if we've got a date for baseline
    If myBaselineDate <> "00:00:00" And myReplannedDate <> "00:00:00" Then
        
        '   Write the dates to the template
        strReturnText = myBaselineDate & "|"
        '   See if the baseline is greater than the re-planned and if so default
        If myBaselineDate > myReplannedDate Then
            strReturnText = strReturnText & myBaselineDate
        Else
            strReturnText = strReturnText & myReplannedDate
        End If
        GetBaselineDate = strReturnText
        Exit Function
    Else
        If myBaselineDate <> "00:00:00" And myReplannedDate = "00:00:00" Then
            strReturnText = myBaselineDate & "|" & myBaselineDate
            GetBaselineDate = strReturnText
            Exit Function
        End If
        If myBaselineDate = "00:00:00" And myReplannedDate <> "00:00:00" Then
            strReturnText = myReplannedDate & "|" & myReplannedDate
            GetBaselineDate = strReturnText
            Exit Function
        End If
        If myActualDate <> "00:00:00" Then
            strReturnText = myActualDate & "|"
            strReturnText = strReturnText & myActualDate
            GetBaselineDate = strReturnText
        Exit Function
        End If
    End If
   
End Function
Public Function GetTestScriptData()
    
    '   Write them out
    strReturn = dblGlobalExecutedPercent & " %|"
    strReturn = strReturn & dblGlobalPlannedPassedPercent & " %|"
    strReturn = strReturn & dblGlobalPlannedFailedPercent & " %|"
    strReturn = strReturn & dblGlobalExecutedPassedPercent & " %|"
    strReturn = strReturn & dblGlobalExecutedFailedPercent & " %"
    
    GetTestScriptData = strReturn
    
End Function
Public Function GetDefectsByPriority()
Dim myUrgentCount As Integer, myHighCount As Integer, myMediumCount As Integer, myLowCount As Integer
Dim myTotal As Integer

    '   Get the defect by priority array
    myPriorityArr = DefectStatusByPriority()
    
    '   Loop round the array getting the different priorities
    For Each Ele In myPriorityArr
        '   Split the element
        mySplit = Split(Ele, "|")
        '   mySplit(1) is Urgent, mySplit(2) is high, mySplit(3) is medium and mySplit(4) is low
        If mySplit(0) = "New" Or mySplit(0) = "Assigned" Or mySplit(0) = "Open" Or mySplit(0) = "Reopen" Or mySplit(0) = "Failed Testing" Then
            myUrgentCount = myUrgentCount + mySplit(1)
            myHighCount = myHighCount + mySplit(2)
            myMediumCount = myMediumCount + mySplit(3)
            myLowCount = myLowCount + mySplit(4)
        End If
    Next
    
    '   Total them all up
    myTotal = myUrgentCount + myHighCount + myMediumCount + myLowCount
    
    '   Write em out
    strReturn = myUrgentCount & "|"
    strReturn = strReturn & myHighCount & "|"
    strReturn = strReturn & myMediumCount & "|"
    strReturn = strReturn & myLowCount & "|"
    strReturn = strReturn & myTotal
    
    '   Return the string
    GetDefectsByPriority = strReturn
    
End Function
Public Function GetDefectsBySeverity()
Dim myCriticalCount As Integer, myHighCount As Integer, myMediumCount As Integer, myLowCount As Integer
Dim myTotal As Integer

    '   Get the defect by priority array
    mySeverityArr = DefectStatusBySeverity()
    
    '   Loop round the array getting the different priorities
    For Each Ele In mySeverityArr
        '   Split the element
        mySplit = Split(Ele, "|")
        '   mySplit(1) is Critical, mySplit(2) is high, mySplit(3) is medium and mySplit(4) is low
        If mySplit(0) = "New" Or mySplit(0) = "Assigned" Or mySplit(0) = "Open" Or mySplit(0) = "Reopen" Or mySplit(0) = "Failed Testing" Then
            myCriticalCount = myCriticalCount + mySplit(1)
            myHighCount = myHighCount + mySplit(2)
            myMediumCount = myMediumCount + mySplit(3)
            myLowCount = myLowCount + mySplit(4)
        End If
    Next
    
    '   Total them all up
    myTotal = myCriticalCount + myHighCount + myMediumCount + myLowCount
    
    '   Write em out
    strReturn = myCriticalCount & "|"
    strReturn = strReturn & myHighCount & "|"
    strReturn = strReturn & myMediumCount & "|"
    strReturn = strReturn & myLowCount & "|"
    strReturn = strReturn & myTotal
    
    '   Return the string
    GetDefectsBySeverity = strReturn
    
End Function
Public Function GetDefectCount(ByVal blnDashboard As Boolean) As Boolean
Dim tdcBugFactory
Dim tdcBugFilter
Dim colBugList
Dim iCount As Integer
iCount = 0

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
        GetDefectCount = False
    Else
        If blnDashboard = True Then
            '   Filter out closed statuses and see if we've still got defects
            For Each objBug In colBugList
                If objBug.Status <> "Closed" And objBug.Status <> "Rejected" And objBug.Status <> "Duplicate" Then
                    iCount = iCount + 1
                End If
            Next
            If iCount = 0 Then
                GetDefectCount = False
            Else
                GetDefectCount = True
            End If
        Else
            GetDefectCount = True
        End If
    End If
    
End Function

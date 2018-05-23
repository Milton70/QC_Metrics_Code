' --------------------------------------------- TD Globals ------------------------------------------------
Public tdc As TDConnection
Public tdcBugFactory As BugFactory
Public tdcTestFactory As TestFactory
Public tdcTestSetFactory As TestSetFactory
Public TSTestFact As TSTestFactory
Public tdcBugFilter As TDFilter
Public tdcTestFilter As TDFilter
Public aField As TDField
Public tdcBug As Bug
Public tdcTest As Test

' --------------------------------------------- Excel Globals ------------------------------------------------
Public objWrkBk As Excel.Workbook
Public objWrkSht As Excel.Worksheet

' --------------------------------------------- Dictionary Globals ------------------------------------------------
Public objDefectAgeBySeverityDictionary As Dictionary
Public objRootCauseBySeverityDictionary As Dictionary
Public objDefectStatusBySeverityDictionary As Dictionary
Public objTestCaseStatusByPriorityDictionary As Dictionary
Public objTestPhaseBySeverityDictionary As Dictionary
Public objBugFunctionalAreaDictionary As Dictionary
Public objTestFunctionalAreaDictionary As Dictionary
Public objTestSetDictionary As Dictionary
Public objDateDictionary As Dictionary
Public objDatesStatusDictionary As Dictionary

' --------------------------------------------- List Globals ------------------------------------------------
Public fieldList As list
Public colBugList As list
Public colTestList As list

' --------------------------------------------- String Globals ------------------------------------------------
Public strTemplatePath As String
Public strUserName As String
Public strPassWord As String
Public strQCProject As String
Public strQCDomain As String
Public strProject As String
Public strProjectName As String
Public strTestPhase As String
Public strSubProjectName As String
Public wrkShtTestSetData As String
Public wrkShtDefectData As String
Public wrkShtScriptProgress As String
Public wrkShtOpenDefects As String
Public wrkShtDefectProgress As String
Public wrkShtDailyDefect As String
Public strField As String
Public strFieldName As String
Public strListName As String
Public strDomain As String
Public strTDProject As String
Public strUser As String
Public strPhaseField As String
Public strPath As String
Public FName As String
Public strHeader As String
Public strTestManager As String
Public strPhaseExStartDate As String
Public strPhaseActStartDate As String
Public strPhaseExEndDate As String
Public strPhaseActEndDate As String
Public strProjectStatus As String
Public strTestStatus As String
Public strRisks As String
Public sPath As String
Public strProjectCycleLabel As String
Public strTestPhaseCycleLabel As String
Public strProjectBugLabel As String
Public strTestPhaseBugLabel As String
Public strSubProjectCycleLabel As String
Public strSubProjectBugLabel As String
Public strRootCauseLabel As String
Public strFixedOnDateLabel As String
Public strTestedOnDateLabel As String
Public strDetectedOnDateLabel As String
Public strClosedOnDateLabel As String
Public strTestCycleLabel As String
Public strFunctionalAreaTestLabel As String
Public strFunctionalAreaBugLabel As String
Public strTestCycle As String
Public strFunctionalArea As String
Public strReplannedDateLabel As String
Public strFolderDate As String
Public strFolderPath As String
Public strPathandFileName As String
Public strWorkbookPath As String
Public strWebPath As String
Public strTheWebPath As String
Public strWebArchivePath As String
Public strDocumentationPath As String

' --------------------------------------------- Date Globals ------------------------------------------------
Public strStartDate As Date
Public strEndDate As Date
Public dateActualStart As Date
Public dateActualEnd As Date
Public datePlannedStart As Date
Public datePlannedEnd As Date
Public dateEndDate As Date
Public dStartDate As Date
Public dtmStartDate As Date
Public dtmEndDate As Date
Public TodaysDate As Date
Public myStartDate As Date
Public myEndDate As Date

' --------------------------------------------- Boolean Globals ------------------------------------------------
Public blnDebug As Boolean
Public blnReportToTemp As Boolean
Public blnRunMetrics As Boolean
Public blnAdd As Boolean
Public blnUpdate As Boolean
Public blnLowNumberDays As Boolean
Public blnFoundForm As Boolean
Public blnUseEarliestDate As Boolean
Public blnEnable As Boolean
Public blnFolderSelected As Boolean
Public blnNewFolderEntered As Boolean
Public blnSelected As Boolean
Public blnWeeklyReports As Boolean
Public blnRunToday As Boolean
Public blnDontRun As Boolean
Public blnGoNoFurther As Boolean
Public blnNoTestSets As Boolean
Public blnAutomationReport As Boolean

' --------------------------------------------- Variant Globals ------------------------------------------------
Public arCycleLabels As Variant
Public arBugLabels As Variant
Public arTestLabels As Variant
Public arTestSets() As Variant
Public arDefects As Variant
Public arrActuals() As Variant
Public arrBaseline() As Variant
Public arrReplanned() As Variant
Public arrPassed() As Variant
Public arrFailed() As Variant
Public arrExecuted() As Variant
Public arrNC() As Variant
Public arrNR() As Variant
Public arrProjects() As Variant
Public arrTestInstance() As Variant
Public arrTR() As Variant
Public arrDummy() As Variant
Public arrTotal() As Variant
Public arrAccepted() As Variant
Public arrFixed() As Variant
Public arrTested() As Variant
Public arrClosed() As Variant

' --------------------------------------------- Integer Globals ------------------------------------------------
Public iPos As Integer
Public intCurrentRow As Integer
Public intTotalPlannedTests As Integer
Public iBase As Integer
Public iRep As Integer
Public iExe As Integer
Public iPass As Integer
Public iFail As Integer
Public iNC As Integer
Public iNR As Integer
Public iTI As Integer
Public iTR As Integer
Public iDum As Integer
Public iFails As Integer

' --------------------------------------------- Double Globals ------------------------------------------------
Public dblGlobalExecutedPercent As Double
Public dblGlobalPlannedPassedPercent As Double
Public dblGlobalPlannedFailedPercent As Double
Public dblGlobalExecutedPassedPercent As Double
Public dblGlobalExecutedFailedPercent As Double

' --------------------------------------------- Misc Globals ------------------------------------------------
Public fso As New FileSystemObject
' --------------------------------------------- This versions Run Metrics ------------------------------------------------
Public Function RunMetrics()
Dim arrProjects() As Variant
Dim iCount As Integer
Dim myTemp As String
myTemp = ""
iCount = -1

    '   Open the projects list spreadsheet and get all the details into an array
    Set objWrkBk = Nothing
    '   Open the projects list
    Set objWrkBk = Workbooks.Open("\\view\general\testing\general_documentation\Daily Status Reports\Current Projects v2.0.xlsm")
        
    '   Set application display alerts off
    objWrkBk.Application.DisplayAlerts = False
    
    '   Get the number of rows in the list
    iTotalRows = objWrkBk.Worksheets(1).UsedRange.Rows.Count
    '   Loop round and put into an array
    For iRow = 2 To iTotalRows
        iCount = iCount + 1
        ReDim Preserve arrProjects(iCount)
        arrProjects(iCount) = objWrkBk.Worksheets(1).Cells(iRow, 6).Value & "|" & objWrkBk.Worksheets(1).Cells(iRow, 7).Value & "|" & objWrkBk.Worksheets(1).Cells(iRow, 1).Value & "|" & objWrkBk.Worksheets(1).Cells(iRow, 2).Value & "|" & objWrkBk.Worksheets(1).Cells(iRow, 3).Value & "|" & objWrkBk.Worksheets(1).Cells(iRow, 4).Value & "|" & objWrkBk.Worksheets(1).Cells(iRow, 5).Value
    Next
    
    '   Close the project sheet
    objWrkBk.Close
    Set objWrkBk = Nothing
    
    '   Convert the date into yyyymmdd format
    strFolderDate = Format(TodaysDate, "yyyymmdd")
    
    '   Create the homepage
    CreateDummyHomepage arrProjects
    
    '   Go through the array and create dashboard files for each project
    For Each ProjectEle In arrProjects
    
        '   Split the array elements into the different parts
        ProjectSplit = Split(ProjectEle, "|")
        
        '   See if we should run this row dependent on it's date range
        If ProjectSplit(0) = "" Then
            GoTo MoveToNext
        End If
        If CDate(ProjectSplit(0)) > TodaysDate Then
            GoTo MoveToNext
        End If
        If ProjectSplit(1) <> "" Then
            If CDate(ProjectSplit(1)) < TodaysDate Then
                GoTo MoveToNext
            End If
        End If
        
        '   Set up qc project and project name variables
        strQCProject = ProjectSplit(2)
        strProjectName = ProjectSplit(3)
        strTestPhase = ProjectSplit(4)
        strSubProjectName = ProjectSplit(5)
        strTestCycle = ProjectSplit(6)
        
        '   Create a new date folder for the QC project
        Select Case strQCProject
            Case "BACK_OFFICE"
                If fso.FolderExists(sPath & "Back Office\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Back Office\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Back Office\Temp\"
                Else
                    strFolderPath = sPath & "Back Office\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Back Office/"
            Case "SHARED_TECHNICAL_SERVICES"
                If fso.FolderExists(sPath & "Shared Technical Services\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Shared Technical Services\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Shared Technical Services\Temp\"
                Else
                    strFolderPath = sPath & "Shared Technical Services\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Shared Technical Services/"
            Case "COMMON_CLEARING_SERVICES"
                If fso.FolderExists(sPath & "Common Clearing Services\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Common Clearing Services\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Common Clearing Services\Temp\"
                Else
                    strFolderPath = sPath & "Common Clearing Services\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Common Clearing Services/"
            Case "EQUITIES"
                If fso.FolderExists(sPath & "Equities\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Equities\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Equities\Temp\"
                Else
                    strFolderPath = sPath & "Equities\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Equities/"
            Case "FIXED_INCOME"
                If fso.FolderExists(sPath & "Fixed Income\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Fixed Income\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Fixed Income\Temp\"
                Else
                    strFolderPath = sPath & "Fixed Income\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Fixed Income/"
            Case "FX"
                If fso.FolderExists(sPath & "FX\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "FX\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "FX\Temp\"
                Else
                    strFolderPath = sPath & "FX\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "FX/"
            Case "GDP"
                If fso.FolderExists(sPath & "GDP\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "GDP\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "GDP\Temp\"
                Else
                    strFolderPath = sPath & "GDP\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "GDP/"
            Case "RISK"
                If fso.FolderExists(sPath & "Risk\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Risk\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Risk\Temp\"
                Else
                    strFolderPath = sPath & "Risk\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Risk/"
            Case "SWAPS"
                If fso.FolderExists(sPath & "Swaps\" & strFolderDate) = False Then
                    fso.CreateFolder sPath & "Swaps\" & strFolderDate
                End If
                If blnReportToTemp = True Then
                    strFolderPath = sPath & "Swaps\Temp\"
                Else
                    strFolderPath = sPath & "Swaps\" & strFolderDate & "\"
                End If
                strTheWebPath = strWebPath & "Swaps/"
        End Select
        
        '   See if the dummy file has been replaced
        If fso.FileExists(strFolderPath & strProjectName & "-dashboardTemp.asp") = False Then
            '   Move the dummy dashboard before we start and replace if run is successful
            fso.CopyFile strTemplatePath & "Dummy-dashboard.asp", strFolderPath & strProjectName & "dummy-dashboard.asp"
        End If
        
        '   Disconnect if already connected
        DisconnectFromQC
    
        '   Login to QC
        LoginToQCProject strUserName, strPassWord
        '   Connect to the project
        ConnectToQC "LCHC_STREAM", strQCProject

        '   Create the QC factories
        CreateFactories
    
        '   Get the QC field labels for these
        strProjectCycleLabel = GetFieldName("Project", "CYCLE")
        strTestPhaseCycleLabel = GetFieldName("Test Phase", "CYCLE")
        strProjectBugLabel = GetFieldName("Project Name", "BUG")
        strTestPhaseBugLabel = GetFieldName("Test Phase", "BUG")
        strTestedOnDateLabel = GetFieldName("Tested On Date", "BUG")
        strClosedOnDateLabel = GetFieldName("Closing Date", "BUG")
        strRootCauseLabel = GetFieldName("Root Cause", "BUG")
        strFixedOnDateLabel = GetFieldName("Fixed On Date", "BUG")
        strDetectedOnDateLabel = "BG_DETECTION_DATE"
        strReplannedDateLabel = GetFieldName("Replanned Exec Date", "TESTCYCL")
    
        '   Sub project stuff
        strSubProjectCycleLabel = GetFieldName("Sub Project", "CYCLE")
        strSubProjectBugLabel = GetFieldName("Sub Project", "BUG")
        
        '   Cycle stuff
        strTestCycleLabel = GetFieldName("Test Cycle", "CYCLE")
        
        '   Functional Area stuff
        strFunctionalAreaTestLabel = GetFieldName("Functional Area", "TEST")
        strFunctionalAreaBugLabel = GetFieldName("Functional Area", "BUG")
    
        '   Set the header to just the project name for the dashboard
        strHeader = strProjectName
        
        '   See if the project name has a period and remove
        strThisProject = Replace(strProjectName, ".", "")
        strThisProject = Replace(strThisProject, ",", "")
        strThisProject = Replace(strThisProject, "+", "")
        strThisProject = Replace(strThisProject, "/", "")
        
        '   See if the sub project name has a period and remove
        strThisSubProject = Replace(strSubProjectName, ".", "")
        strThisSubProject = Replace(strThisSubProject, ",", "")
        strThisSubProject = Replace(strThisSubProject, "+", "")
        strThisSubProject = Replace(strThisSubProject, "/", "")
        
        '   See if the test cycle has a period and remove
        strThisTestCycle = Replace(strTestCycle, ".", "")
        strThisTestCycle = Replace(strThisTestCycle, ",", "")
        strThisTestCycle = Replace(strThisTestCycle, "+", "")
        strThisTestCycle = Replace(strThisTestCycle, "/", "")
        
        '   See if the test cycle has a period and remove
        strThisTestPhase = Replace(strTestPhase, ".", "")
        strThisTestPhase = Replace(strThisTestPhase, ",", "")
        strThisTestPhase = Replace(strThisTestPhase, "+", "")
        strThisTestPhase = Replace(strThisTestPhase, "/", "")
        
        '   Set the path and new filename
        If strSubProjectName <> "N/A" Then
            If strTestCycle <> "N/A" Then
                strPathandFileName = strThisProject & "-" & strThisTestPhase & "-" & strThisSubProject & "-" & strThisTestCycle
            Else
                strPathandFileName = strThisProject & "-" & strThisTestPhase & "-" & strThisSubProject
            End If
        Else
            If strTestCycle <> "N/A" Then
                strPathandFileName = strThisProject & "-" & strThisTestPhase & "-" & strThisTestCycle
            Else
                strPathandFileName = strThisProject & "-" & strThisTestPhase
            End If
        End If
        
        '   Write the project name, test phase and date to the header rows
        If strSubProjectName = "N/A" Then
            If strTestCycle = "N/A" Then
                strHeader = strProjectName & " - " & strTestPhase & " - " & TodaysDate
            Else
                strHeader = strProjectName & " - " & strTestPhase & " - " & strTestCycle & " - " & TodaysDate
            End If
        Else
            If strTestCycle = "N/A" Then
                strHeader = strProjectName & " - " & strTestPhase & " - " & strSubProjectName & " - " & TodaysDate
           Else
                strHeader = strProjectName & " - " & strTestPhase & " - " & strSubProjectName & " - " & strTestCycle & " - " & TodaysDate
            End If
        End If
        
        '   Get Test Data
        blnDontRun = False
        GetTestSets
        If blnDontRun = True Then
            GoTo RunDefects
        End If
        
        '   Get Test Run data
        GetTestsRun
        
        '   Get Tests by Functional Area
        GetTestsByFunctionalArea
        
        '   Get run data
        GetRuns
        
RunDefects:
        blnDontRun = False
        '   Get the defects by severity and priority
        GetDefectsBySeverityByPriority
        If blnDontRun = True Then
            GoTo RunDashboard
        End If
        
        '   Get Daily Defects
        GetDailyDefects
        
        '   Get Defect age data
        GetDefectAgeBySeverity
        
        '   Get defects by functional area
        GetDefectsByFunctionalArea
        
        '   Get Defect Root Cause
        GetDefectRootCause
        
        '   Build Defect details page
        BuildDefectDetails
        
        '   Get Defects by history
        GetDefectsbyHistory
        
        '   Get Defects over time
        GetDefectsOverTime
        
        '   Get Defects open
        GetOpenDefects
         
RunDashboard:
        '   Create the dashboard
        CreateDashboard
        
        '   See if we're running automated tests
        If strTestPhase = "Automation" Then
            blnAutomationReport = True
            RunAutomationReport
        End If
        
        '   Remove all the files we don't need and rename the ones we do
        CleanUp
        
MoveToNext:
    Next
    
    '   Merge all dashboards so we have the correct links
    MergeDashboards arrProjects
    
    '   Create the homepage
    CreateHomepage arrProjects
    
    '   Copy to the web servers
    CopyToWeb

End Function





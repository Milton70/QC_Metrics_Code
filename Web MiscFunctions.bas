Attribute VB_Name = "MiscFunctions"
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const MAX_PATH As Long = 260

Type BrowseInfo
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszINSTRUCTIONS As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Declare Function SHGetPathFromIDListA Lib "shell32.dll" ( _
    ByVal pidl As Long, _
    ByVal pszBuffer As String) As Long

Declare Function SHBrowseForFolderA Lib "shell32.dll" ( _
    lpBrowseInfo As BrowseInfo) As Long
Public Function IsEven(ByVal Number As Long) As Boolean
    IsEven = (Number Mod 2 = 0)
End Function
Function BrowseFolder(Optional Caption As String = "") As String

Dim BrowseInfo As BrowseInfo
Dim FolderName As String
Dim ID As Long
Dim Res As Long

With BrowseInfo
   .hOwner = 0
   .pidlRoot = 0
   .pszDisplayName = String$(MAX_PATH, vbNullChar)
   .lpszINSTRUCTIONS = Caption
   .ulFlags = BIF_RETURNONLYFSDIRS
   .lpfn = 0
End With

FolderName = String$(MAX_PATH, vbNullChar)
ID = SHBrowseForFolderA(BrowseInfo)
If ID Then
   Res = SHGetPathFromIDListA(ID, FolderName)
   If Res Then
       BrowseFolder = Left$(FolderName, InStr(FolderName, vbNullChar) - 1)
   End If
End If

End Function
Public Function FormChecks()
    
    '   See if the combo box values have been selected
    If frmSetUp.cboxProject.Value = "" Then
        MsgBox "Please select a valid project.", vbExclamation
        FormChecks = "Fail"
        Exit Function
    End If
    If frmSetUp.cboxTestPhase.Value = "" Then
        MsgBox "Please select a valid test phase.", vbExclamation
        FormChecks = "Fail"
        Exit Function
    End If
    
End Function
Public Function FindWeekends(ByVal dateStart As Date, ByVal dateEnd As Date)
Dim myArr()
Dim iCount As Integer
iCount = 0
    
    '   Set the array to at least an empty array
    ReDim Preserve myArr(iCount)
    myArr(iCount) = ""
    '   Loop round all dates in range
    For myDate = dateStart To dateEnd
    
        Select Case Weekday(myDate)
            Case vbSaturday
                ReDim Preserve myArr(iCount)
                myArr(iCount) = myDate + 2
                iCount = iCount + 1
        End Select
    Next
    '   See if we've got an array or not
    If myArr(0) = "" Then
        '   We've only got weeks or days to use so return this
        myArr(0) = "No Weeks"
    End If
    
    FindWeekends = myArr
       
End Function
Public Function FindMonths(ByVal dateStart As Date, ByVal dateEnd As Date)
Dim myArr()
Dim iCount As Integer
Dim blnBroke As Boolean
iCount = 0

    '   Set the array to at least an empty array
    ReDim Preserve myArr(iCount)
    myArr(iCount) = ""
    '   Loop round removing months at a time
    Do
        rc = DateAdd("m", 1, dateStart)
        If rc > dateEnd Then
            blnBroke = True
        Else
            ReDim Preserve myArr(iCount)
            myArr(iCount) = rc
            dateStart = rc
            iCount = iCount + 1
        End If
    Loop Until blnBroke = True
    '   Find the number of days from the current start date to the end date
    rc = DateDiff("d", dateStart, dateEnd)
    
    '   See if we've got an array or not
    If myArr(0) = "" Then
        '   We've only got weeks or days to use so return this
        myArr(0) = "No Month"
    Else
        '   See if we're over half way through a month or not
        If rc > 15 Then
            rc = DateAdd("m", 1, dateStart)
            If rc <= dateEnd Then
                ReDim Preserve myArr(iCount)
                myArr(iCount) = rc
            End If
        Else
            '   Add these days to the last end date in the array
            myArr(UBound(myArr)) = DateAdd("d", rc, myArr(UBound(myArr)))
        End If
        '   See if any of our months are weekends
        For i = 0 To UBound(myArr)
        
            If Weekday(myArr(i)) = 1 Then
                myArr(i) = DateAdd("d", 1, myArr(i))
            Else
                If Weekday(myArr(i)) = 7 Then
                    myArr(i) = DateAdd("d", 2, myArr(i))
                End If
            End If
        Next
    End If
    
    FindMonths = myArr

End Function
Public Function FindPreviousCol(ByVal strPrevStatus As String)
    Select Case strPrevStatus
        Case "New"
            iPrevCol = 2
        Case "Assigned"
            iPrevCol = 3
        Case "Open", "In Progress"
            iPrevCol = 4
        Case "Fixed"
            iPrevCol = 5
        Case "Ready For Testing", "Ready for Testing"
            iPrevCol = 6
        Case "Failed Testing"
            iPrevCol = 7
        Case "Tested"
            iPrevCol = 8
        Case "Reopen"
            iPrevCol = 9
        Case "Duplicate"
            iPrevCol = 10
        Case "Rejected"
            iPrevCol = 11
        Case "On Hold"
            iPrevCol = 12
        Case "Closed", "Status_Closed"
            iPrevCol = 13
    End Select
    
    FindPreviousCol = iPrevCol
End Function
Public Function AccumulateData()

    '  Set up the status array
    Dim arrStatus(12)
    arrStatus(0) = "New"
    arrStatus(1) = "Assigned"
    arrStatus(2) = "Open"
    arrStatus(3) = "Fixed"
    arrStatus(4) = "Tested"
    arrStatus(5) = "Ready For Testing"
    arrStatus(6) = "Failed Testing"
    arrStatus(7) = "Reopen"
    arrStatus(8) = "Duplicate"
    arrStatus(9) = "Rejected"
    arrStatus(10) = "On Hold"
    arrStatus(11) = "Closed"
    Dim iPrev
    Dim iCurr
    Dim iTot

    '  Get the dates in the date dictionary into an array
    myDates = objDateDictionary.Keys
    
    '  Loop through the date array
    For i = 0 To UBound(myDates)
        '  If we're on iteration 0 then don't add
        If i <> 0 Then
            '  Loop through the status
            For j = 0 To UBound(arrStatus)
                '  Get the previous value
                iPrev = objDatesStatusDictionary.Item(myDates(i - 1) & "|" & arrStatus(j))
                '  Get the current value for this date
                iCurr = objDatesStatusDictionary.Item(myDates(i) & "|" & arrStatus(j))
                '  Add the two together
                iTot = iPrev + iCurr
                '  Write this new value to the current date
                objDatesStatusDictionary.Item(myDates(i) & "|" & arrStatus(j)) = iTot
            Next
        End If
    Next
    
End Function

Public Function RemoveProjectList()
'   Make the project data worksheet displayed
    ThisWorkbook.Worksheets("Project Data").Visible = True
    
    '   Clear the project data sheet
    ThisWorkbook.Worksheets("Project Data").Activate
    ThisWorkbook.Worksheets("Project Data").Select
    iTotalRows = ThisWorkbook.Worksheets("Project Data").UsedRange.Rows.Count
    If iTotalRows <> 1 Then
        For iRow = 2 To iTotalRows
            ThisWorkbook.Worksheets("Project Data").Rows(iRow & ":" & iRow).Select
            Selection.Delete Shift:=xlUp
            If iRow = 2 Then
                If ThisWorkbook.Worksheets("Project Data").Cells(iRow, 1).Value = "" Then
                    Exit For
                Else
                    iRow = iRow - 1
                End If
            End If
        Next
    End If
    
    '   Make the project data worksheet hidden
    ThisWorkbook.Worksheets("Project Data").Visible = False
End Function
Public Function GetWeekday(ByVal dateTheDate As Date)
    iDay = Weekday(dateTheDate)
    Select Case iDay
        Case 1
            GetWeekday = "Sunday"
        Case 2
            GetWeekday = "Monday"
        Case 3
            GetWeekday = "Tuesday"
        Case 4
            GetWeekday = "Wednesday"
        Case 5
            GetWeekday = "Thursday"
        Case 6
            GetWeekday = "Friday"
        Case 7
            GetWeekday = "Saturday"
    End Select
End Function
Public Function FindLastCol(ByVal strWorksheet As String, ByVal iRow As Integer)
Dim LastCol As Long
  
    With objWrkBk.Worksheets(strWorksheet)
  
        FindLastCol = .Cells(iRow, .Columns.Count).End(xlToLeft).Column
  
    End With

End Function
Public Function FindLastRow(ByVal strWorksheet As String, ByVal iCol As Integer)
Dim LastCol As Long
  
    With objWrkBk.Worksheets(strWorksheet)
  
        FindLastRow = .Cells(.Rows.Count, iCol).End(xlUp).Row
  
    End With

End Function
Public Sub SortDictionary(Dict As Scripting.Dictionary, _
    SortByKey As Boolean, _
    Optional Descending As Boolean = False, _
    Optional CompareMode As VbCompareMethod = vbTextCompare)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SortDictionary
' This sorts a Dictionary object. If SortByKey is False, the
' the sort is done based on the Items of the Dictionary, and
' these items must be simple data types. They may not be
' Object, Arrays, or User-Defined Types. If SortByKey is True,
' the Dictionary is sorted by Key value, and the Items in the
' Dictionary may be Object as well as simple variables.
'
' If sort by key is True, all element of the Dictionary
' must have a non-blank Key value. If Key is vbNullString
' the procedure will terminate.
'
' By defualt, sorting is done in Ascending order. You can
' sort by Descending order by setting the Descending parameter
' to True.
'
' By default, text comparisons are done case-INSENSITIVE (e.g.,
' "a" = "A"). To use case-SENSITIVE comparisons (e.g., "a" <> "A")
' set CompareMode to vbBinaryCompare.
'
' Note: This procedure requires the
' QSortInPlace function, which is described and available for
' download at www.cpearson.com/excel/qsort.htm .
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim KeyValue As String
Dim ItemValue As Variant
Dim Arr() As Variant
Dim KeyArr() As String
Dim VTypes() As VbVarType

Dim V As Variant
Dim SplitArr As Variant

Dim TempDict As Scripting.Dictionary
'''''''''''''''''''''''''''''
' Ensure Dict is not Nothing.
'''''''''''''''''''''''''''''
If Dict Is Nothing Then
    Exit Sub
End If
''''''''''''''''''''''''''''
' If the number of elements
' in Dict is 0 or 1, no
' sorting is required.
''''''''''''''''''''''''''''
If (Dict.Count = 0) Or (Dict.Count = 1) Then
    Exit Sub
End If

''''''''''''''''''''''''''''
' Create a new TempDict.
''''''''''''''''''''''''''''
Set TempDict = New Scripting.Dictionary

If SortByKey = True Then
    ''''''''''''''''''''''''''''''''''''''''
    ' We're sorting by key. Redim the Arr
    ' to the number of elements in the
    ' Dict object, and load that array
    ' with the key names.
    ''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To Dict.Count - 1)
    
    For Ndx = 0 To Dict.Count - 1
        Arr(Ndx) = Dict.Keys(Ndx)
    Next Ndx
    
    ''''''''''''''''''''''''''''''''''''''
    ' Sort the key names.
    ''''''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=CompareMode
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Load TempDict. The key value come from
    ' our sorted array of keys Arr, and the
    ' Item comes from the original Dict object.
    ''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = 0 To Dict.Count - 1
        KeyValue = Arr(Ndx)
        TempDict.Add Key:=KeyValue, Item:=Dict.Item(KeyValue)
    Next Ndx
    '''''''''''''''''''''''''''''''''
    ' Set the passed in Dict object
    ' to our TempDict object.
    '''''''''''''''''''''''''''''''''
    Set Dict = TempDict
    ''''''''''''''''''''''''''''''''
    ' This is the end of processing.
    ''''''''''''''''''''''''''''''''
Else
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Here, we're sorting by items. The Items must
    ' be simple data types. They may NOT be Objects,
    ' arrays, or UserDefineTypes.
    ' First, ReDim Arr and VTypes to the number
    ' of elements in the Dict object. Arr will
    ' hold a string containing
    '   Item & vbNullChar & Key
    ' This keeps the association between the
    ' item and its key.
    '''''''''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To Dict.Count - 1)
    ReDim VTypes(0 To Dict.Count - 1)

    For Ndx = 0 To Dict.Count - 1
        If (IsObject(Dict.Items(Ndx)) = True) Or _
            (IsArray(Dict.Items(Ndx)) = True) Or _
            VarType(Dict.Items(Ndx)) = vbUserDefinedType Then
            Debug.Print "***** ITEM IN DICTIONARY WAS OBJECT OR ARRAY OR UDT"
            Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Here, we create a string containing
        '       Item & vbNullChar & Key
        ' This preserves the associate between an item and its
        ' key. Store the VarType of the Item in the VTypes
        ' array. We'll use these values later to convert
        ' back to the proper data type for Item.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Arr(Ndx) = Dict.Items(Ndx) & vbNullChar & Dict.Keys(Ndx)
            VTypes(Ndx) = VarType(Dict.Items(Ndx))
            
    Next Ndx
    ''''''''''''''''''''''''''''''''''
    ' Sort the array that contains the
    ' items of the Dictionary along
    ' with their associated keys
    ''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=vbTextCompare
    
    For Ndx = LBound(Arr) To UBound(Arr)
        '''''''''''''''''''''''''''''''''''''
        ' Loop trhogh the array of sorted
        ' Items, Split based on vbNullChar
        ' to get the Key from the element
        ' of the array Arr.
        SplitArr = Split(Arr(Ndx), vbNullChar)
        ''''''''''''''''''''''''''''''''''''''''''
        ' It may have been possible that item in
        ' the dictionary contains a vbNullChar.
        ' Therefore, use UBound to get the
        ' key value, which will necessarily
        ' be the last item of SplitArr.
        ' Then Redim Preserve SplitArr
        ' to UBound - 1 to get rid of the
        ' Key element, and use Join
        ' to reassemble to original value
        ' of the Item.
        '''''''''''''''''''''''''''''''''''''''''
        KeyValue = SplitArr(UBound(SplitArr))
        ReDim Preserve SplitArr(LBound(SplitArr) To UBound(SplitArr) - 1)
        ItemValue = Join(SplitArr, vbNullChar)
        '''''''''''''''''''''''''''''''''''''''
        ' Join will set ItemValue to a string
        ' regardless of what the original
        ' data type was. Test the VTypes(Ndx)
        ' value to convert ItemValue back to
        ' the proper data type.
        '''''''''''''''''''''''''''''''''''''''
        Select Case VTypes(Ndx)
            Case vbBoolean
                ItemValue = CBool(ItemValue)
            Case vbByte
                ItemValue = CByte(ItemValue)
            Case vbCurrency
                ItemValue = CCur(ItemValue)
            Case vbDate
                ItemValue = CDate(ItemValue)
            Case vbDecimal
                ItemValue = CDec(ItemValue)
            Case vbDouble
                ItemValue = CDbl(ItemValue)
            Case vbInteger
                ItemValue = CInt(ItemValue)
            Case vbLong
                ItemValue = CLng(ItemValue)
            Case vbSingle
                ItemValue = CLng(ItemValue)
            Case vbString
                ItemValue = CStr(ItemValue)
            Case Else
                ItemValue = ItemValue
        End Select
        ''''''''''''''''''''''''''''''''''''''
        ' Finally, add the Item and Key to
        ' our TempDict dictionary.
        
        TempDict.Add Key:=KeyValue, Item:=ItemValue
    Next Ndx
End If


'''''''''''''''''''''''''''''''''
' Set the passed in Dict object
' to our TempDict object.
'''''''''''''''''''''''''''''''''
Set Dict = TempDict
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modQSortInPlace
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This module contains the QSortInPlace procedure and private supporting procedures.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal LB As Long = -1&, _
    Optional ByVal UB As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortInPlace
'
' This function sorts the array InputArray in place -- this is, the original array in the
' calling procedure is sorted. It will work with either string data or numeric data.
' It need not sort the entire array. You can sort only part of the array by setting the LB and
' UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
' LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
' omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
' omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
' UB to UBound(InputArray).
'
' By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
' To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
' to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
' it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
'
' The function returns TRUE if the array was successfully sorted or FALSE if an error
' occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
' displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
'
''''''''''''''''''''''''''''''''''''''
' MODIFYING THIS CODE:
''''''''''''''''''''''''''''''''''''''
' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
' variable. E.g.,
'       If SomethingThatCausesAnExit Then
'           RecursionLevel = RecursionLevel - 1
'           Exit Function
'       End If
'''''''''''''''''''''''''''''''''''''''
'
' Note: If you coerce InputArray to a ByVal argument, QSortInPlace will not be
' able to reference the InputArray in the calling procedure and the array will
' not be sorted.
'
' This function uses the following procedures. These are declared as Private procedures
' at the end of this module:
'       IsArrayAllocated
'       IsSimpleDataType
'       IsSimpleNumericType
'       QSortCompare
'       NumberOfArrayDimensions
'       ReverseArrayInPlace
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Temp As Variant
Dim Buffer As Variant
Dim CurLow As Long
Dim CurHigh As Long
Dim CurMidpoint As Long
Dim Ndx As Long
Dim pCompareMode As VbCompareMethod

'''''''''''''''''''''''''
' Set the default result.
'''''''''''''''''''''''''
QSortInPlace = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This variable is used to determine the level
' of recursion  (the function calling itself).
' RecursionLevel is incremented when this procedure
' is called, either initially by a calling procedure
' or recursively by itself. The variable is decremented
' when the procedure exits. We do the input parameter
' validation only when RecursionLevel is 1 (when
' the function is called by another function, not
' when it is called recursively).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Static RecursionLevel As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Keep track of the recursion level -- that is, how many
' times the procedure has called itself.
' Carry out the validation routines only when this
' procedure is first called. Don't run the
' validations on a recursive call to the
' procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
RecursionLevel = RecursionLevel + 1

If RecursionLevel = 1 Then
    ''''''''''''''''''''''''''''''''''
    ' Ensure InputArray is an array.
    ''''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' InputArray is not an array. Exit with a False result.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        RecursionLevel = RecursionLevel - 1
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test LB and UB. If < 0 then set to LBound and UBound
    ' of the InputArray.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If LB < 0 Then
        LB = LBound(InputArray)
    End If
    If UB < 0 Then
        UB = UBound(InputArray)
    End If
    
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            ''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is an empty, unallocated array."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case 1
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
        Case Else
            ''''''''''''''''''''''''''''''''''''''''''
            ' We sort ONLY single dimensional arrays.
            ''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The InputArray is multi-dimensional." & _
                      "QSortInPlace works only on single-dimensional arrays."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that InputArray is an array of simple data
    ' types, not other arrays or objects. This tests
    ' the data type of only the first element of
    ' InputArray. If InputArray is an array of Variants,
    ' subsequent data types may not be simple data types
    ' (e.g., they may be objects or other arrays), and
    ' this may cause QSortInPlace to fail on the StrComp
    ' operation.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "InputArray is not an array of simple data types."
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure that the LB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case LB
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is > UB
            If NoAlerts = False Then
                MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ensure the UB parameter is valid.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case UB
        Case Is > UBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is < LBound(InputArray)
            If NoAlerts = False Then
                MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
        Case Is < LB
            If NoAlerts = False Then
                MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
    End Select

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' if UB = LB, we have nothing to sort, so get out.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If UB = LB Then
        QSortInPlace = True
        RecursionLevel = RecursionLevel - 1
        Exit Function
    End If

End If ' RecursionLevel = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ensure that CompareMode is either vbBinaryCompare  or
' vbTextCompare. If it is neither, default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
    pCompareMode = CompareMode
Else
    pCompareMode = vbTextCompare
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Begin the actual sorting process.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CurLow = LB
CurHigh = UB

CurMidpoint = (LB + UB) \ 2 ' note integer division (\) here

Temp = InputArray(CurMidpoint)

Do While (CurLow <= CurHigh)
    
    Do While QSortCompare(V1:=InputArray(CurLow), V2:=Temp, CompareMode:=pCompareMode) < 0
        CurLow = CurLow + 1
        If CurLow = UB Then
            Exit Do
        End If
    Loop
    
    Do While QSortCompare(V1:=Temp, V2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
        CurHigh = CurHigh - 1
        If CurHigh = LB Then
           Exit Do
        End If
    Loop

    If (CurLow <= CurHigh) Then
        Buffer = InputArray(CurLow)
        InputArray(CurLow) = InputArray(CurHigh)
        InputArray(CurHigh) = Buffer
        CurLow = CurLow + 1
        CurHigh = CurHigh - 1
    End If
Loop

If LB < CurHigh Then
    QSortInPlace InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

If CurLow < UB Then
    QSortInPlace InputArray:=InputArray, LB:=CurLow, UB:=UB, _
        Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
End If

'''''''''''''''''''''''''''''''''''''
' If Descending is True, reverse the
' order of the array, but only if the
' recursion level is 1.
'''''''''''''''''''''''''''''''''''''
If Descending = True Then
    If RecursionLevel = 1 Then
        ReverseArrayInPlace InputArray
    End If
End If

RecursionLevel = RecursionLevel - 1
QSortInPlace = True
End Function

Private Function QSortCompare(V1 As Variant, V2 As Variant, _
    Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' QSortCompare
' This function is used in QSortInPlace to compare two elements. If
' V1 AND V2 are both numeric data types (integer, long, single, double)
' they are converted to Doubles and compared. If V1 and V2 are BOTH strings
' that contain numeric data, they are converted to Doubles and compared.
' If either V1 or V2 is a string and does NOT contain numeric data, both
' V1 and V2 are converted to Strings and compared with StrComp.
'
' The result is -1 if V1 < V2,
'                0 if V1 = V2
'                1 if V1 > V2
' For text comparisons, case sensitivity is controlled by CompareMode.
' If this is vbBinaryCompare, the result is case SENSITIVE. If this
' is omitted or any other value, the result is case INSENSITIVE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim D1 As Double
Dim D2 As Double
Dim S1 As String
Dim S2 As String

Dim Compare As VbCompareMethod
''''''''''''''''''''''''''''''''''''''''''''''''
' Test CompareMode. Any value other than
' vbBinaryCompare will default to vbTextCompare.
''''''''''''''''''''''''''''''''''''''''''''''''
If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
    Compare = CompareMode
Else
    Compare = vbTextCompare
End If
'''''''''''''''''''''''''''''''''''''''''''''''
' If either V1 or V2 is either an array or
' an Object, raise a error 13 - Type Mismatch.
'''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V1) = True Or IsArray(V2) = True Then
    Err.Raise 13
    Exit Function
End If
If IsObject(V1) = True Or IsObject(V2) = True Then
    Err.Raise 13
    Exit Function
End If

If IsSimpleNumericType(V1) = True Then
    If IsSimpleNumericType(V2) = True Then
        '''''''''''''''''''''''''''''''''''''
        ' If BOTH V1 and V2 are numeric data
        ' types, then convert to Doubles and
        ' do an arithmetic compare and
        ' return the result.
        '''''''''''''''''''''''''''''''''''''
        D1 = CDbl(V1)
        D2 = CDbl(V2)
        If D1 = D2 Then
            QSortCompare = 0
            Exit Function
        End If
        If D1 < D2 Then
            QSortCompare = -1
            Exit Function
        End If
        If D1 > D2 Then
            QSortCompare = 1
            Exit Function
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''
' Either V1 or V2 was not numeric data type.
' Test whether BOTH V1 AND V2 are numeric
' strings. If BOTH are numeric, convert to
' Doubles and do a arithmetic comparison.
''''''''''''''''''''''''''''''''''''''''''''
If IsNumeric(V1) = True And IsNumeric(V2) = True Then
    D1 = CDbl(V1)
    D2 = CDbl(V2)
    If D1 = D2 Then
        QSortCompare = 0
        Exit Function
    End If
    If D1 < D2 Then
        QSortCompare = -1
        Exit Function
    End If
    If D1 > D2 Then
        QSortCompare = 1
        Exit Function
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''
' Either or both V1 and V2 was not numeric
' string. In this case, convert to Strings
' and use StrComp to compare.
''''''''''''''''''''''''''''''''''''''''''''''
S1 = CStr(V1)
S2 = CStr(V2)
QSortCompare = StrComp(S1, S2, Compare)

End Function



Private Function NumberOfArrayDimensions(Arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Ndx As Integer
Dim Res As Integer
On Error Resume Next
' Loop, increasing the dimension index Ndx, until an error occurs.
' An error will occur when Ndx exceeds the number of dimension
' in the array. Return Ndx - 1.
Do
    Ndx = Ndx + 1
    Res = UBound(Arr, Ndx)
Loop Until Err.Number <> 0

NumberOfArrayDimensions = Ndx - 1

End Function
 
Private Function ReverseArrayInPlace(InputArray As Variant, _
    Optional NoAlerts As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ReverseArrayInPlace
' This procedure reverses the order of an array in place -- this is, the array variable
' in the calling procedure is sorted. An error will occur if InputArray is not an array,
 'if it is an empty, unallocated array, or if the number of dimensions is not 1.
'
' NOTE: Before calling the ReverseArrayInPlace procedure, consider if your needs can
' be met by simply reading the existing array in reverse order (Step -1). If so, you can save
' the overhead added to your application by calling this function.
'
' The function returns TRUE if the array was successfully reversed, or FALSE if
' an error occurred.
'
' If an error occurred, a message box is displayed indicating the error. To suppress
' the message box and simply return FALSE, set the NoAlerts parameter to TRUE.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Temp As Variant
Dim Ndx As Long
Dim Ndx2 As Long

''''''''''''''''''''''''''''''''
' Set the default return value.
''''''''''''''''''''''''''''''''
ReverseArrayInPlace = False

'''''''''''''''''''''''''''''''''
' Ensure we have an array
'''''''''''''''''''''''''''''''''
If IsArray(InputArray) = False Then
    If NoAlerts = False Then
        MsgBox "The InputArray parameter is not an array."
    End If
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''
' Test the number of dimensions of the
' InputArray. If 0, we have an empty,
' unallocated array. Get out with
' an error message. If greater than
' one, we have a multi-dimensional
' array, which is not allowed. Only
' an allocated 1-dimensional array is
' allowed.
''''''''''''''''''''''''''''''''''''''
Select Case NumberOfArrayDimensions(InputArray)
    Case 0
        '''''''''''''''''''''''''''''''''''''''''''
        ' Zero dimensions indicates an unallocated
        ' dynamic array.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
        End If
        Exit Function
    Case 1
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        '''''''''''''''''''''''''''''''''''''''''''
        ' We can reverse ONLY a single dimensional
        ' arrray.
        '''''''''''''''''''''''''''''''''''''''''''
        If NoAlerts = False Then
            MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                   "on single-dimensional arrays."
        End If
        Exit Function

End Select

'''''''''''''''''''''''''''''''''''''''''''''
' Ensure that we have only simple data types,
' not an array of objects or arrays.
'''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
    If NoAlerts = False Then
        MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
            "ReverseArrayInPlace can reverse only arrays of simple data types."
        Exit Function
    End If
End If

Ndx2 = UBound(InputArray)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' loop from the LBound of InputArray to the midpoint of InputArray
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
    '''''''''''''''''''''''''''''''''
    'swap the elements
    '''''''''''''''''''''''''''''''''
    Temp = InputArray(Ndx)
    InputArray(Ndx) = InputArray(Ndx2)
    InputArray(Ndx2) = Temp
    '''''''''''''''''''''''''''''
    ' decrement the upper index
    '''''''''''''''''''''''''''''
    Ndx2 = Ndx2 - 1

Next Ndx
ReverseArrayInPlace = True
End Function

Private Function IsSimpleNumericType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleNumericType
' This returns TRUE if V is one of the following data types:
'        vbBoolean
'        vbByte
'        vbCurrency
'        vbDate
'        vbDecimal
'        vbDouble
'        vbInteger
'        vbLong
'        vbSingle
'        vbVariant if it contains a numeric value
' It returns FALSE for any other data type, including any array
' or vbEmpty.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsSimpleDataType(V) = True Then
    Select Case VarType(V)
        Case vbBoolean, _
                vbByte, _
                vbCurrency, _
                vbDate, _
                vbDecimal, _
                vbDouble, _
                vbInteger, _
                vbLong, _
                vbSingle
            IsSimpleNumericType = True
        Case vbVariant
            If IsNumeric(V) = True Then
                IsSimpleNumericType = True
            Else
                IsSimpleNumericType = False
            End If
        Case Else
            IsSimpleNumericType = False
    End Select
Else
    IsSimpleNumericType = False
End If
End Function

Private Function IsSimpleDataType(V As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsSimpleDataType
' This function returns TRUE if V is one of the following
' variable types (as returned by the VarType function:
'    vbBoolean
'    vbByte
'    vbCurrency
'    vbDate
'    vbDecimal
'    vbDouble
'    vbEmpty
'    vbError
'    vbInteger
'    vbLong
'    vbNull
'    vbSingle
'    vbString
'    vbVariant
'
' It returns FALSE if V is any one of the following variable
' types:
'    vbArray
'    vbDataObject
'    vbObject
'    vbUserDefinedType
'    or if it is an array of any type.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test if V is an array. We can't just use VarType(V) = vbArray
' because the VarType of an array is vbArray + VarType(type
' of array element). E.g, the VarType of an Array of Longs is
' 8195 = vbArray + vbLong.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(V) = True Then
    IsSimpleDataType = False
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We must also explicitly check whether V is an object, rather
' relying on VarType(V) to equal vbObject. The reason is that
' if V is an object and that object has a default proprety, VarType
' returns the data type of the default property. For example, if
' V is an Excel.Range object pointing to cell A1, and A1 contains
' 12345, VarType(V) would return vbDouble, the since Value is
' the default property of an Excel.Range object and the default
' numeric type of Value in Excel is Double. Thus, in order to
' prevent this type of behavior with default properties, we test
' IsObject(V) to see if V is an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsObject(V) = True Then
    IsSimpleDataType = False
    Exit Function
End If
'''''''''''''''''''''''''''''''''''''
' Test the value returned by VarType.
'''''''''''''''''''''''''''''''''''''
Select Case VarType(V)
    Case vbArray, vbDataObject, vbObject, vbUserDefinedType
        '''''''''''''''''''''''
        ' not simple data types
        '''''''''''''''''''''''
        IsSimpleDataType = False
    Case Else
        ''''''''''''''''''''''''''''''''''''
        ' otherwise it is a simple data type
        ''''''''''''''''''''''''''''''''''''
        IsSimpleDataType = True
End Select

End Function

Private Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array has not been allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''
' If Arr is not an array, return FALSE and get out.
'''''''''''''''''''''''''''''''''''''''''''''''''''
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Try to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occured.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
N = UBound(Arr, 1)
If Err.Number = 0 Then
    '''''''''''''''''''''''''''''''''''''
    ' No error. Array has been allocated.
    '''''''''''''''''''''''''''''''''''''
    IsArrayAllocated = True
Else
    '''''''''''''''''''''''''''''''''''''
    ' Error. Unallocated array.
    '''''''''''''''''''''''''''''''''''''
    IsArrayAllocated = False
End If

End Function
Public Function GetDatesAndStatus(dateStart, dateEnd)
Dim blnBroke
Dim rc
blnBroke = "False"

    '   Set up the dictionaries
    Set objDateDictionary = New Dictionary
    Set objDatesStatusDictionary = New Dictionary
    
    '  Add the 1st date
    objDateDictionary.Add dateStart, 0
    objDatesStatusDictionary.Add dateStart & "|New", 0
    objDatesStatusDictionary.Add dateStart & "|Assigned", 0
    objDatesStatusDictionary.Add dateStart & "|Open", 0
    objDatesStatusDictionary.Add dateStart & "|Fixed", 0
    objDatesStatusDictionary.Add dateStart & "|Tested", 0
    objDatesStatusDictionary.Add dateStart & "|Ready For Testing", 0
    objDatesStatusDictionary.Add dateStart & "|Failed Testing", 0
    objDatesStatusDictionary.Add dateStart & "|Reopen", 0
    objDatesStatusDictionary.Add dateStart & "|Duplicate", 0
    objDatesStatusDictionary.Add dateStart & "|Rejected", 0
    objDatesStatusDictionary.Add dateStart & "|On Hold", 0
    objDatesStatusDictionary.Add dateStart & "|Closed", 0

    '   Loop round adding days at a time
    Do

        myDateStart = DateAdd("d", 1, dateStart)
    
        If myDateStart > dateEnd Then
                blnBroke = "True"
        Else
            '  Add the date and every status with a zero value
            objDateDictionary.Add myDateStart, 0
            objDatesStatusDictionary.Add myDateStart & "|New", 0
            objDatesStatusDictionary.Add myDateStart & "|Assigned", 0
            objDatesStatusDictionary.Add myDateStart & "|Open", 0
            objDatesStatusDictionary.Add myDateStart & "|Fixed", 0
            objDatesStatusDictionary.Add myDateStart & "|Tested", 0
            objDatesStatusDictionary.Add myDateStart & "|Ready For Testing", 0
            objDatesStatusDictionary.Add myDateStart & "|Failed Testing", 0
            objDatesStatusDictionary.Add myDateStart & "|Reopen", 0
            objDatesStatusDictionary.Add myDateStart & "|Duplicate", 0
            objDatesStatusDictionary.Add myDateStart & "|Rejected", 0
            objDatesStatusDictionary.Add myDateStart & "|On Hold", 0
            objDatesStatusDictionary.Add myDateStart & "|Closed", 0
        End If
        dateStart = myDateStart
    Loop While blnBroke = "False"
    
End Function
Public Function CleanUp()
    
    '   Rename the ones we want
    If fso.FileExists(strFolderPath & strPathandFileName & "-DefectStatus.asp") = True Then
        fso.DeleteFile strFolderPath & strPathandFileName & "-DefectStatus.asp", True
    End If
    fso.MoveFile strFolderPath & strPathandFileName & "-DefectsStatusStage1.txt", strFolderPath & strPathandFileName & "-DefectStatus.asp"
    
    If fso.FileExists(strFolderPath & strPathandFileName & "-TestSetStatus.asp") = True Then
        fso.DeleteFile strFolderPath & strPathandFileName & "-TestSetStatus.asp", True
    End If
    fso.MoveFile strFolderPath & strPathandFileName & "-TestSetStatusStage1.txt", strFolderPath & strPathandFileName & "-TestSetStatus.asp"
    
    '   Add the links to the test and defect files
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatus.asp", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
    strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
    strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
    strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
    strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
    strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
    strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
    strNewText = Replace(strNewText, "<a href=TestSetLink>Test Status</a> |", "")
    
    '   See if we're running automation
    If blnAutomationReport = True Then
        strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
    Else
        strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
    End If
    
    strNewText = Replace(strNewText, "TablePlannedvsExecuted", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestPlannedvsExecutedTable.asp" & Chr(39))
    
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetStatus.asp", ForWriting, True)
    myFile.WriteLine strNewText
    myFile.Close
    
    '   Defect status
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectStatus.asp", ForReading)
    strText = myFile.ReadAll
    myFile.Close
    
    strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
    strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
    strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
    strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
    strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
    strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
    strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
    strNewText = Replace(strNewText, "<a href=DefectStatusLink>Defect Status</a> |", "")
    strNewText = Replace(strNewText, "TableDetectedAndClosed", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsDetectedvsClosedTable.asp" & Chr(39))
    strNewText = Replace(strNewText, "TableDefectsTime", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsOverTimeTable.asp" & Chr(39))
    
    '   See if we're running automation
    If blnAutomationReport = True Then
        strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
    Else
        strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
    End If
    
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectStatus.asp", ForWriting, True)
    myFile.WriteLine strNewText
    myFile.Close
    
    '   Defects by history
    If fso.FileExists(strFolderPath & strPathandFileName & "-DefectsByHistory.asp") = False Then
        fso.CopyFile strFolderPath & strPathandFileName & "-DefectStatus.asp", strFolderPath & strPathandFileName & "-DefectsByHistory.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        strNewText = Replace(strText, "DefectsByHistory", "DefectStatus")
        strNewText = Replace(strNewText, "Defect History", "Defect Status")
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    Else
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "<a href=DefectsByHistoryLink>Defect History</a> |", "")
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectsByHistory.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Open Defects
    If fso.FileExists(strFolderPath & strPathandFileName & "-OpenDefects.asp") = False Then
        fso.CopyFile strFolderPath & strPathandFileName & "-DefectStatus.asp", strFolderPath & strPathandFileName & "-OpenDefects.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-OpenDefects.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        strNewText = Replace(strText, "OpenDefects", "DefectStatus")
        strNewText = Replace(strNewText, "Open Defects", "Defect Status")
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-OpenDefects.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    Else
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-OpenDefects.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        
        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "<a href=OpenDefectsLink>Outstanding Defects</a> |", "")
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-OpenDefects.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Test Details
    If fso.FileExists(strFolderPath & strPathandFileName & "-TestSetDetails.asp") = False Then
        fso.CopyFile strFolderPath & strPathandFileName & "-TestSetStatus.asp", strFolderPath & strPathandFileName & "-TestSetDetails.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        strNewText = Replace(strText, "TestSetDetails", "TestSetStatus")
        strNewText = Replace(strNewText, "Supporting Test Set Details", "Test Script Status")
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    Else
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
    
        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "FunctionalAreaTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestFunctionalAreaTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "TableTestRuns", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestRunsTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "<a href=TestSetDetailsLink>Supporting Test Details</a> |", "")

        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If

        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-TestSetDetails.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Defect Details
    If fso.FileExists(strFolderPath & strPathandFileName & "-DefectDetails.asp") = False Then
        fso.CopyFile strFolderPath & strPathandFileName & "-DefectStatus.asp", strFolderPath & strPathandFileName & "-DefectDetails.asp"
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDetails.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
        strNewText = Replace(strText, "DefectDetails", "DefectStatus")
        strNewText = Replace(strNewText, "Supporting Defect Details", "Defect Status")
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDetails.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    Else
        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDetails.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close
    
        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", "Supporting Defect Details for " & strHeader)
        strNewText = Replace(strNewText, "NewFixedTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectNewFixedTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "FixedTestedTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectFixedTestedTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "TimeOpenTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectTimeOpenTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "RootCauseTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectRootCauseTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "FunctionalAreaTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectFunctionalAreaTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDaily", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDaily.aspx" & Chr(39))
        strNewText = Replace(strNewText, "DDTable", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDailyTable.asp" & Chr(39))
        strNewText = Replace(strNewText, "<a href=DefectDetailsLink>Supporting Defect Details</a> |", "")

        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If

        Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDetails.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Set the links on the table files
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectNewFixedTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectNewFixedTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectNewFixedTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If

    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectDailyTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectDailyTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectDailyTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If

    '   Set the links on the table files
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectFixedTestedTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectFixedTestedTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectFixedTestedTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If

    '   Set the links on the table files
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectTimeOpenTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectTimeOpenTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
         
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
    
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectTimeOpenTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Set the links on the table files
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectRootCauseTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectRootCauseTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectRootCauseTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Set the links on the table files
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectFunctionalAreaTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectFunctionalAreaTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectFunctionalAreaTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    '   Set the links on the table files
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-TestFunctionalAreaTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-TestFunctionalAreaTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-TestFunctionalAreaTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-TestPlannedvsExecutedTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-TestPlannedvsExecutedTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-TestPlannedvsExecutedTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectsDetectedvsClosedTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectsDetectedvsClosedTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectsDetectedvsClosedTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-DefectsOverTimeTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectsOverTimeTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-DefectsOverTimeTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-TestRunsTable.asp") = True Then
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-TestRunsTable.asp", ForReading)
        strText = myFile.ReadAll
        myFile.Close

        strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
        strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
        strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
        strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
        strNewText = Replace(strNewText, "strHeader", strHeader)
        
        '   See if we're running automation
        If blnAutomationReport = True Then
            strNewText = Replace(strNewText, "AutomationLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-AutomationReport.asp" & Chr(39))
        Else
            strNewText = Replace(strNewText, "<a href=AutomationLink>Automation Failure Report</a> |", "")
        End If
        
        Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-TestRunsTable.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
    End If
    
    If fso.FileExists(strFolderPath & "/" & strPathandFileName & "-AutomationReport.asp") = True Then
       Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-AutomationReport.asp", ForReading)
       strText = myFile.ReadAll
       myFile.Close

       strNewText = Replace(strText, "DashboardLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(39))
       strNewText = Replace(strNewText, "TestSetLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp" & Chr(39))
       strNewText = Replace(strNewText, "DefectStatusLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectStatus.asp" & Chr(39))
       strNewText = Replace(strNewText, "DefectsByHistoryLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectsByHistory.asp" & Chr(39))
       strNewText = Replace(strNewText, "OpenDefectsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-OpenDefects.asp" & Chr(39))
       strNewText = Replace(strNewText, "TestSetDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetDetails.asp" & Chr(39))
       strNewText = Replace(strNewText, "DefectDetailsLink", Chr(39) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectDetails.asp" & Chr(39))
       strNewText = Replace(strNewText, "BackToHomePageLink", Chr(39) & strWebPath & "default.asp" & Chr(39))
       strNewText = Replace(strNewText, "strHeader", strHeader)
        
       Set myFile = fso.OpenTextFile(strFolderPath & "/" & strPathandFileName & "-AutomationReport.asp", ForWriting, True)
       myFile.WriteLine strNewText
       myFile.Close
    End If
    
End Function
Public Function CreateDashboard()
    
    '   See if the 2nd part of the dashboard already exist (which means we're running more than one project in this stream)
    If fso.FileExists(strFolderPath & strProjectName & "-dashboard.asp") = False Then
        '   Copy the 2nd part of the dashboard to the correct place with the correct name
        fso.CopyFile strTemplatePath & "DashboardTemplate2.txt", strFolderPath & strProjectName & "-dashboard.asp"
    End If
    
    '   Get Baseline Completion date
    strContinue = GetBaselineDate
    '   See if we continue with this test phase
    If strContinue <> "False" Then
    
        '   Move the current data to the write string
        strWriteString = strContinue
    
        '   Get Test Script Data
        strContinue = GetTestScriptData
        
        '   Move the current data to the write string
        strWriteString = strWriteString & "|" & strContinue
    Else
        strWriteString = "&nbsp;|&nbsp;|0|0|0|0|0"
    End If
    '   See if we've got any Defect information for the criteria
    blnContinue = GetDefectCount(True)
    '   See if we get defect data
    If blnContinue = True Then
    
        '   See if the test sets stuff is false
        If strContinue = "False" Then
            strWriteString = "&nbsp;|&nbsp;|&nbsp;|&nbsp;|&nbsp;|&nbsp;|&nbsp;"
        End If
        
        '   Get Defects by Priority
        strContinue = GetDefectsByPriority
        
        '   Move the current data to the write string
        strWriteString = strWriteString & "|" & strContinue

        '   Get Defects by Severity
        strContinue = GetDefectsBySeverity
        
        '   Move the current data to the write string
        strWriteString = strWriteString & "|" & strContinue
    Else
        '   zeros for defects
        strWriteString = strWriteString & "|0|0|0|0|0|0|0|0|0|0"
    End If
    
    '   Write out the row to the html dashboard and colour this row
    Set myFile = fso.OpenTextFile(strFolderPath & strProjectName & "-dashboard.asp", ForAppending, True)
    mySplit = Split(strWriteString, "|")
    '   See if we've got a sub project or not
    If strSubProjectName <> "N/A" Then
        If strTestCycle <> "N/A" Then
            myFile.WriteLine "<tr align=center><td align=left><a href='" & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp'>" & strTestPhase & " - " & strSubProjectName & " - " & strTestCycle & "</a></td><td bgcolor=#C5E3BF>" & mySplit(0) & "</td><td bgcolor=#C5E3BF>" & mySplit(1) & "</td><td bgcolor=#F0FFF0>" & mySplit(2) & "</td><td bgcolor=#F0FFF0>" & mySplit(3) & "</td><td bgcolor=#F0FFF0>" & mySplit(4) & "</td><td bgcolor=#F0FFF0>" & mySplit(5) & "</td><td bgcolor=#F0FFF0>" & mySplit(6) & "</td><td bgcolor=#F6C9CC>" & mySplit(7) & "</td><td bgcolor=#F6C9CC>" & mySplit(8) & "</td><td bgcolor=#F6C9CC>" & mySplit(9) & "</td><td bgcolor=#F6C9CC>" & mySplit(10) & "</td><td bgcolor=#F6C9CC>" & mySplit(11) & "</td><td bgcolor=#F0FFFF>" & mySplit(12) & "</td><td  bgcolor=#F0FFFF>" & mySplit(13) & "</td><td bgcolor=#F0FFFF>" & mySplit(14) & "</td><td bgcolor=#F0FFFF>" & mySplit(15) & "</td><td bgcolor=#F0FFFF>" & mySplit(16) & "</td></tr>"
        Else
            myFile.WriteLine "<tr align=center><td align=left><a href='" & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp'>" & strTestPhase & " - " & strSubProjectName & "</a></td><td bgcolor=#C5E3BF>" & mySplit(0) & "</td><td bgcolor=#C5E3BF>" & mySplit(1) & "</td><td bgcolor=#F0FFF0>" & mySplit(2) & "</td><td bgcolor=#F0FFF0>" & mySplit(3) & "</td><td bgcolor=#F0FFF0>" & mySplit(4) & "</td><td bgcolor=#F0FFF0>" & mySplit(5) & "</td><td bgcolor=#F0FFF0>" & mySplit(6) & "</td><td bgcolor=#F6C9CC>" & mySplit(7) & "</td><td bgcolor=#F6C9CC>" & mySplit(8) & "</td><td bgcolor=#F6C9CC>" & mySplit(9) & "</td><td bgcolor=#F6C9CC>" & mySplit(10) & "</td><td bgcolor=#F6C9CC>" & mySplit(11) & "</td><td bgcolor=#F0FFFF>" & mySplit(12) & "</td><td  bgcolor=#F0FFFF>" & mySplit(13) & "</td><td bgcolor=#F0FFFF>" & mySplit(14) & "</td><td bgcolor=#F0FFFF>" & mySplit(15) & "</td><td bgcolor=#F0FFFF>" & mySplit(16) & "</td></tr>"
        End If
    Else
        If strTestCycle <> "N/A" Then
            myFile.WriteLine "<tr align=center><td align=left><a href='" & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp'>" & strTestPhase & " - " & strTestCycle & "</a></td><td bgcolor=#C5E3BF>" & mySplit(0) & "</td><td bgcolor=#C5E3BF>" & mySplit(1) & "</td><td bgcolor=#F0FFF0>" & mySplit(2) & "</td><td bgcolor=#F0FFF0>" & mySplit(3) & "</td><td bgcolor=#F0FFF0>" & mySplit(4) & "</td><td bgcolor=#F0FFF0>" & mySplit(5) & "</td><td bgcolor=#F0FFF0>" & mySplit(6) & "</td><td bgcolor=#F6C9CC>" & mySplit(7) & "</td><td bgcolor=#F6C9CC>" & mySplit(8) & "</td><td bgcolor=#F6C9CC>" & mySplit(9) & "</td><td bgcolor=#F6C9CC>" & mySplit(10) & "</td><td bgcolor=#F6C9CC>" & mySplit(11) & "</td><td bgcolor=#F0FFFF>" & mySplit(12) & "</td><td  bgcolor=#F0FFFF>" & mySplit(13) & "</td><td bgcolor=#F0FFFF>" & mySplit(14) & "</td><td bgcolor=#F0FFFF>" & mySplit(15) & "</td><td bgcolor=#F0FFFF>" & mySplit(16) & "</td></tr>"
        Else
            myFile.WriteLine "<tr align=center><td align=left><a href='" & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-TestSetStatus.asp'>" & strTestPhase & "</a></td><td bgcolor=#C5E3BF>" & mySplit(0) & "</td><td bgcolor=#C5E3BF>" & mySplit(1) & "</td><td bgcolor=#F0FFF0>" & mySplit(2) & "</td><td bgcolor=#F0FFF0>" & mySplit(3) & "</td><td bgcolor=#F0FFF0>" & mySplit(4) & "</td><td bgcolor=#F0FFF0>" & mySplit(5) & "</td><td bgcolor=#F0FFF0>" & mySplit(6) & "</td><td bgcolor=#F6C9CC>" & mySplit(7) & "</td><td bgcolor=#F6C9CC>" & mySplit(8) & "</td><td bgcolor=#F6C9CC>" & mySplit(9) & "</td><td bgcolor=#F6C9CC>" & mySplit(10) & "</td><td bgcolor=#F6C9CC>" & mySplit(11) & "</td><td bgcolor=#F0FFFF>" & mySplit(12) & "</td><td  bgcolor=#F0FFFF>" & mySplit(13) & "</td><td bgcolor=#F0FFFF>" & mySplit(14) & "</td><td bgcolor=#F0FFFF>" & mySplit(15) & "</td><td bgcolor=#F0FFFF>" & mySplit(16) & "</td></tr>"
        End If
    End If
    myFile.Close
    
End Function
Public Function CreateHomepage(ByRef arrProjects As Variant)
Dim myArr() As Variant
    '   Copy the homepage template
    fso.CopyFile strTemplatePath & "MetricsHomepageTemplate.txt", sPath & "default.asp"
    
    '   Open the homepage
    Set myFile = fso.OpenTextFile(sPath & "default.asp", ForAppending, True)

    '   Loop round the projects
    For i = 0 To UBound(arrProjects)
        
        '   Split the array elements into the different parts
        ProjectSplit = Split(arrProjects(i), "|")
        strProjectName = ProjectSplit(3)
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
        If myTemp = strProjectName Then
            GoTo MoveToNext
        End If
        If myTemp = "" Or myTemp <> strProjectName Then
            myTemp = strProjectName
            If myQCTemp = "" Or myQCTemp <> ProjectSplit(2) Then
                myQCTemp = ProjectSplit(2)
                blnFolderCopied = False
                blnProjectNodeCreated = False
            Else
                blnFolderCopied = True
                blnProjectNodeCreated = True
            End If
        End If
        
        '   Set up qc project and project name variables
        strQCProject = ProjectSplit(2)
        '   Convert to the formatted version
        Select Case strQCProject
            Case "BACK_OFFICE"
                ThisProject = "Back Office"
                strTheWebPath = strWebPath & "Back Office/"
            Case "SHARED_TECHNICAL_SERVICES"
                ThisProject = "Shared Technical Services"
                strTheWebPath = strWebPath & "Shared Technical Services/"
            Case "COMMON_CLEARING_SERVICES"
                ThisProject = "Common Clearing Services"
                strTheWebPath = strWebPath & "Common Clearing Services/"
            Case "EQUITIES"
                ThisProject = "Equities"
                strTheWebPath = strWebPath & "Equities/"
            Case "FIXED_INCOME"
                ThisProject = "Fixed Income"
                strTheWebPath = strWebPath & "Fixed Income/"
	    Case "FX"
                ThisProject = "FX"
                strTheWebPath = strWebPath & "FX/"	
            Case "GDP"
                ThisProject = "Synapse"
                strTheWebPath = strWebPath & "GDP/"
            Case "RISK"
                ThisProject = "Risk"
                strTheWebPath = strWebPath & "Risk/"
            Case "SWAPS"
                ThisProject = "Swaps"
                strTheWebPath = strWebPath & "Swaps/"
        End Select
        
        '   Get project name
        strProjectName = ProjectSplit(3)

        '   Write the project details
        If blnProjectNodeCreated = False Then
            myFile.WriteLine ("<li>")
            myFile.WriteLine ("<A onmouseover=" & Chr(34) & "this.style.cursor='hand'" & Chr(34) & "onClick=" & Chr(34) & "Toggle(this)" & Chr(34) & "><u> " & ThisProject & "</u></A><DIV style='display:none'>")
            myFile.WriteLine ("<a href=" & Chr(34) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(34) & "> " & strProjectName & "</a><DIV style='display:none'></DIV></li>")
            blnProjectNodeCreated = True
        Else
            myFile.WriteLine ("</br>")
            myFile.WriteLine ("<a href=" & Chr(34) & strTheWebPath & strFolderDate & "/" & strProjectName & "-dashboard.asp" & Chr(34) & "> " & strProjectName & "</a><DIV style='display:none'></DIV>")
        End If
        '   See if the next array element is the same QC project or not
        If i < UBound(arrProjects) Then
            rc = arrProjects(i + 1)
            myrcSplit = Split(rc, "|")
            If myrcSplit(2) <> ProjectSplit(2) Then
                myFile.WriteLine ("</DIV></li>")
            End If
        Else
            myFile.WriteLine ("</DIV></li>")
        End If
MoveToNext:
    Next
    '   Write out the remainder
    myFile.WriteLine ("</ul>")
    myFile.WriteLine ("If you would like to view any of the last 10 archived reports, please select the date from the list and click on Go.")
    myFile.WriteLine ("</br>")
    myFile.WriteLine ("<b>Archived Metrics Reports</b>")
    myFile.WriteLine ("</br></br>")
    myFile.WriteLine ("<form name=" & Chr(34) & "dropdown" & Chr(34) & ">")
    myFile.WriteLine ("<select name=" & Chr(34) & "list" & Chr(34) & ">")
    
    '   Go through the archive and pull out all the dates
    Set myFolder = fso.GetFolder(sPath & "Archive\")
    i = -1
    For Each theFile In myFolder.Files
        i = i + 1
        ReDim Preserve myArr(i)
        myArr(i) = Left(theFile.Name, 8)
    Next
    '   Sort the array
    BubbleSort myArr
    '   Loop round putting only last 10 into file
    For i = UBound(myArr) To UBound(myArr) - 9 Step -1
        myFile.WriteLine ("<option value=" & Chr(34) & strWebArchivePath & myArr(i) & "_default.asp" & Chr(34) & ">" & myArr(i) & "</option>")
    Next
    
    myFile.WriteLine ("</select>")
    myFile.WriteLine ("<input type=button value=" & Chr(34) & "Go" & Chr(34) & " onclick=" & Chr(34) & "goToNewPage(document.dropdown.list)" & Chr(34) & ">")
    myFile.WriteLine ("</form>")
    myFile.WriteLine ("To view documentation on the web based QC metrics, please click the link below:")
    myFile.WriteLine ("</br></br>")
    myFile.WriteLine ("<a href=" & Chr(34) & strDocumentationPath & Chr(34) & ">User Guide</a>")
    myFile.WriteLine ("</td>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("</table>")
    myFile.WriteLine ("<%")
    myFile.WriteLine ("FinishPage();")
    myFile.WriteLine ("%>")
    myFile.WriteLine ("</body>")
    myFile.WriteLine ("</html>")
    myFile.Close
    
    '   Set up the correct message for the side pane
    Set myFile = fso.OpenTextFile(sPath & "default.asp", ForReading, True)
    strText = myFile.ReadAll
    myFile.Close
    
    strText = Replace(strText, "strHeader", Chr(34) & "QC Metrics for " & TodaysDate & Chr(34))
    
    Set myFile = fso.OpenTextFile(sPath & "default.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close
    
End Function
Public Function CreateDummyHomepage(ByRef arrProjects As Variant)
Dim myArr() As Variant
Dim mySourcePath As String
Dim myTargetPath1 As String
Dim myTargetPath2 As String
Dim myDefaultPageSourcePath As String
Dim myDefaultPageTargetPath1 As String
Dim myDefaultPageTargetPath2 As String
Dim rc As String
Dim mySourceFolder As Folder

    '   Get the weekday to write to the header
    strWeekDay = GetWeekday(TodaysDate)

    '   If it exists, move the current homepage into the archive and give a date prefix
    If fso.FileExists(sPath & "default.asp") = True Then
        '   See what day we're on and get previous weekday
        If Weekday(TodaysDate) = 2 Then
            OurDate = Format(DateAdd("d", -3, TodaysDate), "yyyymmdd")
        Else
            OurDate = Format(DateAdd("d", -1, TodaysDate), "yyyymmdd")
        End If
        
        '       Open the file for reading
        Set mySource = fso.OpenTextFile(sPath & "default.asp", ForReading)
        '       Open the file for writing
        Set myDest = fso.OpenTextFile(sPath & "Archive\" & OurDate & "_default.asp", ForWriting, True)

        Do While mySource.AtEndOfStream <> True
                rc = mySource.ReadLine
                If InStr(1, rc, "<form name=" & Chr(34) & "dropdown" & Chr(34) & ">") > 0 Then
                        myDest.WriteLine "<ul><li><a href=" & Chr(34) & strWebPath & Chr(34) & ">Back to Homepage</a></li></ul>"
                Else
                        If InStr(1, rc, "<select name=" & Chr(34) & "list" & Chr(34) & ">") > 0 Then
                        Else
                                If InStr(1, rc, "<option value=") > 0 Then
                                Else
                                        If InStr(1, rc, "</select") > 0 Then
                                        Else
                                                If InStr(1, rc, "<input type=button") > 0 Then
                                                Else
                                                        If InStr(1, rc, "</form>") > 0 Then
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

        '   Archive old stuff
        'ArchiveOffOldFiles

    End If
    
    '   Copy the homepage template
    fso.CopyFile strTemplatePath & "MetricsHomepageTemplate.txt", sPath & "default.asp"
    
    '   Open the homepage
    Set myFile = fso.OpenTextFile(sPath & "default.asp", ForAppending, True)

    '   Loop round the projects
    For i = 0 To UBound(arrProjects)
        
        '   Split the array elements into the different parts
        ProjectSplit = Split(arrProjects(i), "|")
        strProjectName = ProjectSplit(3)
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
        If myTemp = strProjectName Then
            GoTo MoveToNext
        End If
        If myTemp = "" Or myTemp <> strProjectName Then
            myTemp = strProjectName
            If myQCTemp = "" Or myQCTemp <> ProjectSplit(2) Then
                myQCTemp = ProjectSplit(2)
                blnFolderCopied = False
                blnProjectNodeCreated = False
            Else
                blnFolderCopied = True
                blnProjectNodeCreated = True
            End If
        End If
        
        '   Set up qc project and project name variables
        strQCProject = ProjectSplit(2)
        '   Convert to the formatted version
        Select Case strQCProject
            Case "BACK_OFFICE"
                ThisProject = "Back Office"
                strTheWebPath = strWebPath & "Back Office/"
            Case "SHARED_TECHNICAL_SERVICES"
                ThisProject = "Shared Technical Services"
                strTheWebPath = strWebPath & "Shared Technical Services/"
            Case "COMMON_CLEARING_SERVICES"
                ThisProject = "Common Clearing Services"
                strTheWebPath = strWebPath & "Common Clearing Services/"
            Case "EQUITIES"
                ThisProject = "Equities"
                strTheWebPath = strWebPath & "Equities/"
            Case "FIXED_INCOME"
                ThisProject = "Fixed Income"
                strTheWebPath = strWebPath & "Fixed Income/"
	    Case "FX"
                ThisProject = "FX"
                strTheWebPath = strWebPath & "FX/"
            Case "GDP"
                ThisProject = "Synapse"
                strTheWebPath = strWebPath & "GDP/"
            Case "RISK"
                ThisProject = "Risk"
                strTheWebPath = strWebPath & "Risk/"
            Case "SWAPS"
                ThisProject = "Swaps"
                strTheWebPath = strWebPath & "Swaps/"
        End Select
        
        '   Get project name
        strProjectName = ProjectSplit(3)

        '   Write the project details
        If blnProjectNodeCreated = False Then
            myFile.WriteLine ("<li>")
            myFile.WriteLine ("<A onmouseover=" & Chr(34) & "this.style.cursor='hand'" & Chr(34) & "onClick=" & Chr(34) & "Toggle(this)" & Chr(34) & "><u> " & ThisProject & "</u></A><DIV style='display:none'>")
            myFile.WriteLine ("<a href=" & Chr(34) & strTheWebPath & strFolderDate & "/" & strProjectName & "dummy-dashboard.asp" & Chr(34) & "> " & strProjectName & "</a><DIV style='display:none'></DIV>")
            blnProjectNodeCreated = True
        Else
            myFile.WriteLine ("</br>")
            myFile.WriteLine ("<a href=" & Chr(34) & strTheWebPath & strFolderDate & "/" & strProjectName & "dummy-dashboard.asp" & Chr(34) & "> " & strProjectName & "</a><DIV style='display:none'></DIV>")
        End If
        '   See if the next array element is the same QC project or not
        If i < UBound(arrProjects) Then
            rc = arrProjects(i + 1)
            myrcSplit = Split(rc, "|")
            If myrcSplit(2) <> ProjectSplit(2) Then
                myFile.WriteLine ("</DIV></li>")
            End If
        Else
            myFile.WriteLine ("</DIV></li>")
        End If
MoveToNext:
    Next
    '   Write out the remainder
    myFile.WriteLine ("</ul>")
    myFile.WriteLine ("</td>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("</td>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("<tr>")
    myFile.WriteLine ("<th>Archived Metrics Reports</th>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("<tr>")
    myFile.WriteLine ("<td>")
    myFile.WriteLine ("<form name=" & Chr(34) & "dropdown" & Chr(34) & ">")
    myFile.WriteLine ("<select name=" & Chr(34) & "list" & Chr(34) & ">")
    
    '   Go through the archive and pull out all the dates
    Set myFolder = fso.GetFolder(sPath & "Archive\")
    i = -1
    For Each theFile In myFolder.Files
        i = i + 1
        ReDim Preserve myArr(i)
        myArr(i) = Left(theFile.Name, 8)
    Next
    '   Sort the array
    BubbleSort myArr
    '   Loop round putting only last 10 into file
    For i = UBound(myArr) To UBound(myArr) - 9 Step -1
        myFile.WriteLine ("<option value=" & Chr(34) & strWebArchivePath & myArr(i) & "_default.asp" & Chr(34) & ">" & myArr(i) & "</option>")
    Next

    myFile.WriteLine ("</select>")
    myFile.WriteLine ("<input type=button value=" & Chr(34) & "Go" & Chr(34) & " onclick=" & Chr(34) & "goToNewPage(document.dropdown.list)" & Chr(34) & ">")
    myFile.WriteLine ("</form>")
    myFile.WriteLine ("</td>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("<tr>")
    myFile.WriteLine ("<th>Metrics Documentation</th>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("<tr>")
    myFile.WriteLine ("<td><ul><li><a href=" & Chr(34) & strDocumentationPath & Chr(34) & ">User Guide</a></li></ul>")
    myFile.WriteLine ("</td>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("</td>")
    myFile.WriteLine ("</tr>")
    myFile.WriteLine ("<%")
    myFile.WriteLine ("FinishPage();")
    myFile.WriteLine ("%>")
    myFile.WriteLine ("</body>")
    myFile.WriteLine ("</html>")
    myFile.Close
    
    '   Now move over to web servers

    '   Set the source path
    'mySourcePath = sPath
    
    '   Set the target paths
    'myTargetPath1 = "\\intpr1\qc$\"
    'myTargetPath2 = "\\intpr2\qc$\"
    
    '   Set the default page paths
    'myDefaultPageSourcePath = mySourcePath & "default.asp"
    'myDefaultPageTargetPath1 = myTargetPath1 & "default.asp"
    'myDefaultPageTargetPath2 = myTargetPath2 & "default.asp"
    
    '   Copy the source path default.asp to the target paths
    'fso.CopyFile myDefaultPageSourcePath, myDefaultPageTargetPath1, True
    'fso.CopyFile myDefaultPageSourcePath, myDefaultPageTargetPath2, True


End Function
Public Function MergeDashboards(ByRef arrProjects As Variant)
myTemp = ""
    '   Go through all projects in the list
    For Each ProjectEle In arrProjects
    
        '   Split the array elements into the different parts
        ProjectSplit = Split(ProjectEle, "|")
        
        '   Set variable names
        strQCProject = ProjectSplit(2)
        strProjectName = ProjectSplit(3)
        strTestPhase = ProjectSplit(4)
        strSubProjectName = ProjectSplit(5)
        strTestCycle = ProjectSplit(6)
        
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
        If myTemp = strProjectName Then
            GoTo MoveToNext
        End If
        If myTemp = "" Or myTemp <> strProjectName Then
            myTemp = strProjectName
        End If
        
        '   Convert to the formatted version
        Select Case strQCProject
            Case "BACK_OFFICE"
                ThisProject = "Back Office"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "SHARED_TECHNICAL_SERVICES"
                ThisProject = "Shared Technical Services"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "COMMON_CLEARING_SERVICES	"
                ThisProject = "Common Clearing Services"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "EQUITIES"
                ThisProject = "Equities"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "FIXED_INCOME"
                ThisProject = "Fixed Income"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
	    Case "FX"
                ThisProject = "FX"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "GDP"
                ThisProject = "GDP"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "RISK"
                ThisProject = "Risk"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
            Case "SWAPS"
                ThisProject = "Swaps"
                strFolderPath = sPath & ThisProject & "\" & strFolderDate & "\"
        End Select
        
        '   Open the file
        Set myFile = fso.OpenTextFile(strFolderPath & strProjectName & "-dashboard.asp", ForReading)
        strTempText = myFile.ReadAll
        myFile.Close
        
        '   Replace the dashboard link
        OurDate = Format(CDate(TodaysDate), "dd mmm yy")
        strText = Replace(strTempText, "BackToHomepageLink", Chr(34) & strWebPath & "default.asp" & Chr(34))
        strNewText = Replace(strText, "strHeader", "Dashboard for " & strProjectName & " - " & OurDate)
        strNewText = Replace(strNewText, "ReplaceProjectName", strProjectName)
        'myFile.Close
        
        '   Write the value file
        'Set myFile = fso.OpenTextFile(strFolderPath & strProjectName & "-dashboard.asp", ForAppending, True)
        Set myFile = fso.OpenTextFile(strFolderPath & strProjectName & "-dashboard.asp", ForWriting, True)
        myFile.WriteLine strNewText
        myFile.Close
        
        '   Add remaining code to close the dashboard
        Set myFile = fso.OpenTextFile(strFolderPath & strProjectName & "-dashboard.asp", ForAppending, True)
        myFile.WriteLine "</table>"
        myFile.WriteLine "<%FinishPage();%>"
        myFile.WriteLine "</body></html>"
        myFile.Close
        
MoveToNext:
    Next
End Function
Public Function BuildDefectDetails()

    '   Open the defect details template
    Set myFile = fso.OpenTextFile(strTemplatePath & "DefectsDetailsTemplate.txt", ForReading)
    strTempText = myFile.ReadAll
    myFile.Close
    
    '   Replace the holders with the real graphs
    strText = Replace(strTempText, "DefectNewFixed", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectNewFixed.aspx" & Chr(34))
    strText = Replace(strText, "DefectFixedTested", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectFixedTested.aspx" & Chr(34))
    strText = Replace(strText, "DefectTimeOpen", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectTimeOpen.aspx" & Chr(34))
    strText = Replace(strText, "DefectRootCause", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectRootCause.aspx" & Chr(34))
    strText = Replace(strText, "DefectFunctionalArea", Chr(34) & strTheWebPath & strFolderDate & "/" & strPathandFileName & "-DefectFunctionalArea.aspx" & Chr(34))

    '   Create the actual file and write the correct data
    Set myFile = fso.OpenTextFile(strFolderPath & strPathandFileName & "-DefectDetails.asp", ForWriting, True)
    myFile.WriteLine strText
    myFile.Close

End Function
Public Function CopyToWeb()

Dim mySourcePath As String
Dim myTargetPath1 As String
Dim myTargetPath2 As String
Dim myDefaultPageSourcePath As String
Dim myDefaultPageTargetPath1 As String
Dim myDefaultPageTargetPath2 As String
Dim rc As String
Dim mySourceFolder As Folder

    '   Set the source path
    mySourcePath = sPath
    
    '   Set the target paths
    myTargetPath1 = "\\intpr1\qc$\"
    myTargetPath2 = "\\intpr2\qc$\"
    
    '   Set the default page paths
    myDefaultPageSourcePath = mySourcePath & "default.asp"
    myDefaultPageTargetPath1 = myTargetPath1 & "default.asp"
    myDefaultPageTargetPath2 = myTargetPath2 & "default.asp"
    
    '   Copy the source path default.asp to the target paths
    fso.CopyFile myDefaultPageSourcePath, myDefaultPageTargetPath1, True
    fso.CopyFile myDefaultPageSourcePath, myDefaultPageTargetPath2, True

    '   Work through folders in the source folder
    Set mySourceFolder = fso.GetFolder(mySourcePath)
    Set mySourceFolders = mySourceFolder.SubFolders
    For Each myFolder In mySourceFolders
        If myFolder.Name <> "images" And myFolder.Name <> "Templates" Then
            '   See if we're dealing with the Archive
            If myFolder.Name = "Archive" Then
                rc = Format(DateAdd("d", -1, TodaysDate), "yyyymmdd")
                For Each myFile In myFolder.Files
                    If InStr(1, myFile.Name, rc) > 0 Then
                        '   Copy this to the target archives
                        fso.CopyFile myFile, myTargetPath1 & "Archive\", True
                        fso.CopyFile myFile, myTargetPath2 & "Archive\", True
                    End If
                Next
            Else
                '   See if there are sub folders
                If myFolder.SubFolders.Count > 0 Then
                    '   Set the subfolder name
                    mySubFolderName = myFolder.Name
                    '   If there are, see if one matches our date
                    Set mySourceSubFolders = myFolder.SubFolders
                    For Each mySubFolder In mySourceSubFolders
                        If mySubFolder.Name = strFolderDate Then
                            myCopyPath = mySubFolder.ParentFolder.Path
                            fso.CopyFolder myCopyPath & "\" & strFolderDate, myTargetPath1 & mySubFolderName & "\", True
                            fso.CopyFolder myCopyPath & "\" & strFolderDate, myTargetPath2 & mySubFolderName & "\", True
                        End If
                    Next
                End If
            End If
        End If
    Next

End Function
Public Function CopyToWebReRun()

Dim mySourcePath As String
Dim myTargetPath1 As String
Dim myTargetPath2 As String

    '   Set the source path
    mySourcePath = sPath
    
    '   Set the target paths
    myTargetPath1 = "\\intpr1\qc$\"
    myTargetPath2 = "\\intpr2\qc$\"
    
    '   Copy across the files from the Temp directory to the correct date directory on the web servers
    fso.CopyFile mySourcePath & strProject & "\Temp\*", myTargetPath1 & strProject & "\" & strFolderDate, True
    fso.CopyFile mySourcePath & strProject & "\Temp\*", myTargetPath2 & strProject & "\" & strFolderDate, True
    

End Function
Sub BubbleSort(ToSort As Variant, Optional SortAscending As Boolean = True)
    ' Chris Rae's VBA Code Archive - http://chrisrae.com/vba
    ' By Chris Rae, 19/5/99. My thanks to
    ' Will Rickards and Roemer Lievaart
    ' for some fixes.
    Dim AnyChanges As Boolean
    Dim BubbleSort As Long
    Dim SwapFH As Variant
    Do
        AnyChanges = False
        For BubbleSort = LBound(ToSort) To UBound(ToSort) - 1
            If (ToSort(BubbleSort) > ToSort(BubbleSort + 1) And SortAscending) _
               Or (ToSort(BubbleSort) < ToSort(BubbleSort + 1) And Not SortAscending) Then
                ' These two need to be swapped
                SwapFH = ToSort(BubbleSort)
                ToSort(BubbleSort) = ToSort(BubbleSort + 1)
                ToSort(BubbleSort + 1) = SwapFH
                AnyChanges = True
            End If
        Next BubbleSort
    Loop Until Not AnyChanges
End Sub
Public Function ArchiveOffOldFiles()
Dim strSourcePath1 As String
Dim strSourcePath2 As String
Dim strsourcePath3 As String
Dim myArr() As Variant
Dim i As Integer
i = -1

    '   Don't care if it errors
    On Error Resume Next

    '   Set the source paths
    strSourcePath1 = "\\intpr1\qc$\"
    strSourcePath2 = "\\intpr2\qc$\"
    'strsourcePath3 = "\\view\general\TEST_DISCIPLINE\4.Facilities\Web Based Metrics\Web Runs\"
    strsourcePath3 = sPath

    '   Start with the Archive
    myPath1 = strSourcePath1 & "Archive\"
    myPath2 = strSourcePath2 & "Archive\"
    
    '   Get the number of files in this folder
    Set myFolder = fso.GetFolder(myPath1)
    iNoFiles = myFolder.Files.Count
    '   If 10 or more get rid of oldest ones until only 9
    If iNoFiles >= 10 Then
        '   Put them all into an array after stripping the date out
        For Each sFile In myFolder.Files
            i = i + 1
            ReDim Preserve myArr(i)
            myArr(i) = Left(sFile.Name, 8)
        Next
        '   Sort the array
        BubbleSort myArr
        '   Get the upper boundary of the array
        iBound = UBound(myArr)
        '   If it's more than 9, find out by how much
        iRes = iBound - 9
        If iRes = 0 Then
            '   Just delete the first one in the list
            fso.DeleteFile myPath1 & myArr(0) & "_default.asp", True
            fso.DeleteFile myPath2 & myArr(0) & "_default.asp", True
        Else
            For i = 0 To iRes - 1
                fso.DeleteFile myPath1 & myArr(i) & "_default.asp", True
                fso.DeleteFile myPath2 & myArr(i) & "_default.asp", True
            Next
        End If
       
    End If
    
    '   Now go through the folders and remove the last lot - they are held on clearcase anyway.
    Set mySourceFolder = fso.GetFolder(strSourcePath1)
    Set mySourceFolders = mySourceFolder.SubFolders
    For Each myFolder In mySourceFolders
        i = -1
        If myFolder.Name <> "images" And myFolder.Name <> "Templates" And myFolder.Name <> "Archive" Then
            '   See if there are sub folders
            If myFolder.SubFolders.Count > 0 Then
                '   Set the subfolder name
                mySubFolderName = myFolder.Name
                '   If there are, see if one matches our date
                Set mySourceSubFolders = myFolder.SubFolders
                For Each mySubFolder In mySourceSubFolders
                    i = i + 1
                    ReDim Preserve myArr(i)
                    myArr(i) = Left(mySubFolder.Name, 8)
                Next
                '   Sort the array
                BubbleSort myArr
                '   Get the upper boundary of the array
                iBound = UBound(myArr)
                '   If it's more than 10, find out by how much
                iRes = iBound - 11
                If iRes = 0 Then
                    '   Just delete the first one in the list
                    fso.DeleteFolder strSourcePath1 & mySubFolderName & "\" & myArr(0), True
                    fso.DeleteFolder strSourcePath2 & mySubFolderName & "\" & myArr(0), True
                Else
                    For i = 0 To iRes
                        fso.DeleteFolder strSourcePath1 & mySubFolderName & "\" & myArr(i), True
                        fso.DeleteFolder strSourcePath2 & mySubFolderName & "\" & myArr(i), True
                    Next
                End If
            End If
        End If
    Next

    '   Put Errors back
    On Error GoTo 0
End Function

Attribute VB_Name = "QCFunctions"
Public Function IsConnected() As Boolean

If tdc Is Nothing Then
    IsConnected = False
Else
    IsConnected = tdc.Connected
End If

End Function
Public Sub DisconnectFromQC()

If IsConnected Then
    tdc.Disconnect
    tdc.ReleaseConnection
    ' Destroy the object
    Set tdc = Nothing
End If

End Sub
Public Function ConnectToQC(strDomain As String, strTDProject As String)
'Method to connect to Quaility Centre.  This takes the domain, product,
'username and password as input strings

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If
    
    '   Connect to the domain and project
    tdc.Connect strDomain, strTDProject
    
Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function LoginToQCProject(strUser As String, strPassWord As String)
'Method to login to QC Project.  Takes project and password as inputs.

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    '   Exit out if we're already logged in
    Set tdc = New TDAPIOLELib.TDConnection
    If tdc.Connected = True Then
        If tdc.LoggedIn Then
            Exit Function
        End If
    End If
    tdc.InitConnectionEx "http://cmutility:8080/qcbin"
    
    tdc.Login strUser, strPassWord
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
    
End Function
Public Sub CreateFactories()
Dim tdcBugFactory
Dim tdcBugFilter
Dim tdcTestSetFactory
Dim tdcTestFactory

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

'Set up the Bug Factory
Set tdcBugFactory = tdc.BugFactory
Set tdcBugFilter = tdcBugFactory.Filter
'Set up the Test Set Factory
Set tdcTestSetFactory = tdc.TestSetFactory
'Set up the Test Factory
Set tdcTestFactory = tdc.TestFactory

Exit Sub

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Sub
Public Function GetListValues(strListName As String) As list

'This function gets the values in a QC list
If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    Dim custom As Customization
    Dim customLists As CustomizationLists
    Dim customList As CustomizationList
    Dim Node As CustomizationListNode
    Dim lstValues As list
    Dim strValue As Variant
    
    Dim i As Integer
    
    'Get the cust list
    Set custom = tdc.Customization
    Set customLists = custom.lists
    Set customList = customLists.list(strListName)

    'Get the tree node that represents the list
    Set Node = customList.RootNode
    
    Set lstValues = New list

    'Loop through list and assign each value to array.
    i = 1
    Do Until i = Node.ChildrenCount + 1
        strValue = Node.Children(i).Name
        lstValues.Add (strValue)
        i = i + 1
    Loop
    
    Set GetListValues = lstValues
    
    Exit Function
    
ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
  
End Function
Public Function GetFieldLabels(strTableName As String) As Variant
'This function builds an array contain all the label names of fields in an entity

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

Dim i As Integer
Dim arLabels()

    '   Get the list of fields
    Set fieldList = tdc.Fields(strTableName)
    
    'Loop through all of the fields in the entity and assign the label name to an array
    i = -1
    For Each myField In fieldList
        i = i + 1
        ReDim Preserve arLabels(i)
        arLabels(i) = myField.Property.UserLabel
    Next
    
    'Return the array of labels
    GetFieldLabels = arLabels
    
    Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
    
End Function
Public Function GetFieldName(strField As String, strTable As String) As String
'This function searches an array of labels to return a field name
Dim fieldList

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

Dim i As Integer
    
    '   Set up the field list
    Set fieldList = tdc.Fields(strTable)
    
    'Loop through the array of field labels to find the one that matches the input field
    For i = 1 To fieldList.Count
        If fieldList.Item(i).Property.UserLabel = strField Then
            GetFieldName = fieldList.Item(i).Name
            Exit Function
        End If
    Next
    
Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function
Public Function customLists(strFieldName As String, strTable As String) As String
'This function finds the list which is associated with a field

If blnDebug = False Then
    On Error GoTo ErrorHandler
End If

    Dim cust As Customization
    Dim custFields As CustomizationFields
    Dim aCustField As CustomizationField
    Dim custLists As CustomizationLists
    Dim aCustList As CustomizationList
    Dim listName As String
    
    Set cust = tdc.Customization
    Set custFields = cust.Fields
    
    Set aCustField = custFields.Field(strTable, strFieldName)
    
    Set aCustList = aCustField.list
    listName = aCustList.Name

    customLists = listName
    
Exit Function

ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next
    
End Function
Sub ErrorHandler(MyErr As ErrObject)
     
    Dim Prompt          As String
    Dim Title           As String
    Dim MyResponse      As VbMsgBoxResult

    '   Show the app again
    Application.Visible = True
    '   Create the message
    Prompt = "The following error has occured:" & vbCrLf & vbCrLf & MyErr.Description & " - Ending now"
    Title = "Error"
     
    '   Display the message and clean up
    MyResponse = MsgBox(Prompt, vbOKOnly, Title)
    If objWrkBk Is Nothing Then
    Else
        objWrkBk.Save
        objWrkBk.Close
        Set objWrkSht = Nothing
        Set objWrkBk = Nothing
    End If
    DisconnectFromQC
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    End
        
    On Error GoTo 0
     
End Sub

Public Function GetComboValues(strListName As String) As Variant
'This function gets the values in a QC list

On Error GoTo ErrorHandler

    Dim custom As Customization
    Dim lists As CustomizationLists
    Dim list As CustomizationList
    Dim Node As CustomizationListNode
    
    Dim i As Integer
    
    'Get the cust list
    Set custom = tdc.Customization
    Set lists = custom.lists
    Set list = lists.list(strListName)

    'Get the tree node that represents the list
    Set Node = list.RootNode
    ReDim arListValues(Node.ChildrenCount) As String

    'Loop through list and assign each value to array.
    i = 1
    Do Until i > Node.ChildrenCount
        arListValues(i - 1) = Node.Children(i).Name
        i = i + 1
    Loop
    
    'Return the array of list values
    GetComboValues = arListValues
    
    Exit Function
    
ErrorHandler:
    Call ErrorHandler(Err)
    Resume Next

End Function

















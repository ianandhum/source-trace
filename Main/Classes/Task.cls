VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTaskId As Integer 'local copy
Private mvarTaskName As String 'local copy
Private mvarDescription As String 'local copy
Private mvarStartDate As Date 'local copy
Private mvarEndDate As Date 'local copy
Private mvarCFClass As String 'local copy
Private mvarTags As String 'local copy
Private mvarWeight As Double 'local copy
Private mvarPreNode As Integer 'local copy
Private mvarStatus As String 'local copy
Private mvarIsLoaded As Boolean 'local copy
Private mvarIsDirty As Boolean 'local copy
Private mvarIsNew As Boolean 'local copy
Private mvarIsDeleted As Boolean 'local copy

Private arrayData() As String
Private arrayCols() As String
Private arrayDataTypes() As Integer

Private mvarProjectId As Integer 'local copy

Private mvarCardName As String 'local copy

Public Property Get IsNew() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsNew
    IsNew = mvarIsNew
End Property


Public Property Get IsDeleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsDeleted
    IsDeleted = mvarIsDeleted
End Property


Private Sub Class_Initialize()
    setDataDefinitions
End Sub


Public Sub initTaskWithRecord(taskRecord As Recordset)
    Dim errStatus As Boolean
    errStatus = False
    If taskRecord.State = adStateOpen And Not taskRecord.EOF Then
        If Not IsNull(taskRecord!task_id) Then
            mvarTaskId = taskRecord!task_id
        Else
            errStatus = True
        End If
        
        If Not IsNull(taskRecord!name) Then
            mvarTaskName = taskRecord!name
        Else
            errStatus = True
        End If
        
        If Not IsNull(taskRecord!Description) Then
            mvarDescription = taskRecord!Description
        End If
        
        If Not IsNull(taskRecord!start_date) Then
            mvarStartDate = taskRecord!start_date
        Else
            errStatus = True
        End If
        
        If Not IsNull(taskRecord!end_date) Then
            mvarEndDate = taskRecord!end_date
        Else
            errStatus = True
        End If
        
        If Not IsNull(taskRecord!cf_class) Then
            mvarCFClass = taskRecord!cf_class
        End If
        
        If Not IsNull(taskRecord!Tags) Then
            mvarTags = taskRecord!Tags
        End If
        
        If Not IsNull(taskRecord!Weight) Then
            mvarWeight = taskRecord!Weight
        Else
            errStatus = True
        End If
        
        If Not IsNull(taskRecord!pre_node) Then
            mvarPreNode = taskRecord!pre_node
        End If
        
        If Not IsNull(taskRecord!Status) Then
            mvarStatus = taskRecord!Status
        Else
            errStatus = True
        End If
        
        If Not IsNull(taskRecord!card_name) Then
            mvarCardName = taskRecord!card_name
        Else
            errStatus = True
        End If
        If Not IsNull(taskRecord!project_id) Then
            mvarProjectId = taskRecord!project_id
        Else
            errStatus = True
        End If
    Else
        errStatus = True
    End If
    mvarIsLoaded = Not errStatus
End Sub

Public Sub debugTask(Optional currTaskName As String = "Task")
    Debug.Print "-----------------------"
    Debug.Print "Debugging Task :: " & currTaskName
    Debug.Print "TaskId :" & mvarTaskId
    Debug.Print "TaskName :" & mvarTaskName
    Debug.Print "TaskDesc :" & mvarDescription
    Debug.Print "TaskStartDate :" & mvarStartDate
    Debug.Print "TaskEndDate :" & mvarEndDate
    Debug.Print "TaskClass :" & mvarCFClass
    Debug.Print "TaskTags :" & mvarTags
    Debug.Print "TaskWeight :" & mvarWeight
    Debug.Print "TaskPreNode :" & mvarPreNode
    Debug.Print "TaskStatus :" & mvarStatus
    Debug.Print "TaskIsLoaded :" & mvarIsLoaded
    Debug.Print "TaskIsDirty :" & mvarIsDirty
    Debug.Print "TaskIsNew :" & mvarIsNew
    Debug.Print "ProjectId :" & mvarProjectId
    Debug.Print "CardName :" & mvarCardName
    Debug.Print "-----------------------"
End Sub



Public Function AddNew() As Integer
    mvarIsNew = True
End Function



Public Function Delete() As Integer
    mvarIsDeleted = True
End Function


Public Function SaveChanges() As Integer
    SerializeForDB
    If mvarIsDirty And mvarIsLoaded Then
        SaveChanges = updateDB(arrayCols, arrayData, arrayDataTypes, dbTaskTable, "task_id = " & mvarTaskId)
    ElseIf mvarIsDirty And mvarIsNew Then
        SaveChanges = insertToDB(arrayCols, arrayData, arrayDataTypes, dbTaskTable)
        Dim rec As New Recordset
        Set rec = runQuery("SELECT MAX(task_id) as last_id from tb_task")
        If Not rec.EOF Then
            mvarTaskId = rec!last_id
        End If
    ElseIf mvarIsDeleted And mvarIsLoaded Then
        SaveChanges = deleteFromDB(dbTaskTable, "task_id = " & mvarTaskId)
    Else
        SaveChanges = -1
    End If
End Function


'Convert Variables to array
Private Sub SerializeForDB()
    arrayData(0) = mvarTaskName
    arrayData(1) = mvarDescription
    arrayData(2) = mvarStartDate
    arrayData(3) = mvarEndDate
    arrayData(4) = getNullIfEmpty(mvarCFClass)
    arrayData(5) = mvarTags
    arrayData(6) = mvarWeight
    arrayData(7) = mvarPreNode
    arrayData(8) = mvarStatus
    arrayData(9) = mvarCardName
    arrayData(10) = mvarProjectId
End Sub

Private Sub setDataDefinitions()

    ReDim arrayCols(11)
    arrayCols(0) = "name"
    arrayCols(1) = "description"
    arrayCols(2) = "start_date"
    arrayCols(3) = "end_date"
    arrayCols(4) = "cf_class"
    arrayCols(5) = "tags"
    arrayCols(6) = "weight"
    arrayCols(7) = "pre_node"
    arrayCols(8) = "status"
    arrayCols(9) = "card_name"
    arrayCols(10) = "project_id"

    ReDim arrayDataTypes(11)
    arrayDataTypes(0) = adVarChar
    arrayDataTypes(1) = adVarChar
    arrayDataTypes(2) = adDate
    arrayDataTypes(3) = adDate
    arrayDataTypes(4) = adVarChar
    arrayDataTypes(5) = adVarChar
    arrayDataTypes(6) = adDouble
    arrayDataTypes(7) = adInteger
    arrayDataTypes(8) = adVarChar
    arrayDataTypes(9) = adVarChar
    arrayDataTypes(10) = adInteger
    
    ReDim arrayData(11)
    
End Sub


Private Function getNullIfEmpty(Data As String)
    Dim result As String
    If Data = "" Then
        result = "NULL"
    Else
        result = Data
    End If
    getNullIfEmpty = result
End Function

Public Function LoadSingleton(TaskId As Integer, Optional loadAssocFields As Boolean = True) As Boolean
    If DbConnection.State = adStateOpen Then
        Dim tmpStore As Recordset
        Set tmpStore = selectFromDB(" [SourceTrace].[dbo].[tb_task] ", " task_id = " & TaskId)
        If querySuccess Then initTaskWithRecord tmpStore
    End If
End Function

Public Sub GetProgress()
    
End Sub


Public Property Get IsDirty() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsDirty
    IsDirty = mvarIsDirty
End Property


Public Property Get IsLoaded() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsLoaded
    IsLoaded = mvarIsLoaded
End Property



Public Property Let Status(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
    mvarIsDirty = True
End Property


Public Property Get Status() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
End Property



Public Property Let PreNode(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PreNode = 5
    mvarPreNode = vData
    mvarIsDirty = True
End Property


Public Property Get PreNode() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PreNode
    PreNode = mvarPreNode
End Property



Public Property Let Weight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vData
    mvarIsDirty = True
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let Tags(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tags = 5
    mvarTags = vData
    mvarIsDirty = True
End Property


Public Property Get Tags() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tags
    Tags = mvarTags
End Property



Public Property Let CFClass(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CFClass = 5
    mvarCFClass = vData
    mvarIsDirty = True
End Property


Public Property Get CFClass() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CFClass
    CFClass = mvarCFClass
End Property



Public Property Let EndDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EndDate = 5
    mvarEndDate = vData
    mvarIsDirty = True
End Property


Public Property Get EndDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndDate
    EndDate = mvarEndDate
End Property



Public Property Let StartDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartDate = 5
    mvarStartDate = vData
    mvarIsDirty = True
End Property


Public Property Get StartDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartDate
    StartDate = mvarStartDate
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
    mvarIsDirty = True
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let TaskName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TaskName = 5
    mvarTaskName = vData
    mvarIsDirty = True
End Property


Public Property Get TaskName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaskName
    TaskName = mvarTaskName
End Property


Public Property Get TaskId() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TaskId
    TaskId = mvarTaskId
End Property

Public Property Let cardName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CardName = 5
    mvarCardName = vData
    mvarIsDirty = True
End Property


Public Property Get cardName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CardName
    cardName = mvarCardName
End Property



Public Property Let ProjectId(ByVal vData As Integer)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProjectId = Form1
    mvarProjectId = vData
End Property


Public Property Get ProjectId() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProjectId
    ProjectId = mvarProjectId
    mvarIsDirty = True
End Property

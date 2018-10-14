VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarTasks As Collection 'local copy
Private mvarIsLoaded As Boolean 'local copy
Private mvarTaskCount As Integer 'local copy
Private mvarIsDebug As Integer 'local copy


Private Sub Class_Initialize()
    mvarTaskCount = 0
End Sub

Public Function SaveChanges() As Integer
    Dim updateCount As Integer
    updateCount = 0
    If mvarIsLoaded And Not mvarTasks Is Nothing Then
        For i = 1 To mvarTasks.Count
            updateCount = updateCount + mvarTasks.Item(i).SaveChanges()
        Next i
    Else
        updateCount = -1
    End If
    SaveChanges = updateCount
End Function

Public Sub loadTasksFromDB(Optional condition As String = " 1=1 ")
    
    Dim taskRecords As Recordset
    Set taskRecords = selectFromDB(dbTaskTable, condition)
    If querySuccess Then
        Set mvarTasks = Nothing
        Set mvarTasks = New Collection
        Dim tmpTask As New Task
        While Not taskRecords.EOF
            tmpTask.initTaskWithRecord taskRecords
            If mvarIsDebug Then tmpTask.debugTask "Displaying Tasks"
            mvarTasks.Add tmpTask
            taskRecords.MoveNext
        Wend
        mvarTaskCount = mvarTasks.Count
        mvarIsLoaded = True
    Else
        mvarIsLoaded = False
    End If
End Sub

Public Sub removeTask(taskId As Integer)
    Dim tmpTask As New Task
    tmpTask.LoadSingleton (taskId)
    If tmpTask.IsLoaded Then
        tmpTask.Delete
        tmpTask.SaveChanges
    End If
End Sub

Public Sub updateTaskInfo(updatedTask As Task)
    updatedTask.SaveChanges
End Sub

Public Function addTask(newTask As Task) As Integer
    If Not newTask Is Nothing Then
       addTask = newTask.SaveChanges
    Else
        addTask = -1
    End If
End Function

Public Function getTaskById(ByVal taskId As Integer) As Task
    Dim tmpTask As New Task
    tmpTask.LoadSingleton taskId
    Set getTaskById = tmpTask
    
End Function

Public Function getTasks() As Collection
    If Not mvarIsLoaded Then
        loadTasksFromDB
    End If
    Set getTasks = mvarTasks
    
End Function


Public Property Get Tasks() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tasks
    If Not mvarIsLoaded Then
        loadTasksFromDB
    End If
    Set Tasks = mvarTasks
End Property

Public Property Get IsLoaded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsLoaded
    IsLoaded = mvarIsLoaded
End Property


Public Property Get IsDebug() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsLoaded
    IsDebug = mvarIsDebug
End Property

Public Property Let IsDebug(ByVal vData As Boolean)
    mvarIsDebug = vData
End Property

Public Property Get Count() As Integer
    Count = mvarTaskCount
End Property



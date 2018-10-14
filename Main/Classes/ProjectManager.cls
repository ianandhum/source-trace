VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable(s) to hold property value(s)
Private mvarProjects As Collection 'local copy
Private mvarIsLoaded As Boolean 'local copy
Private mvarProjectCount As Integer 'local copy
Private mvarIsDebug As Integer 'local copy


Private Sub Class_Initialize()
    mvarProjectCount = 0
End Sub

Public Function SaveChanges() As Integer
    Dim updateCount As Integer
    updateCount = 0
    If mvarIsLoaded And Not mvarProjects Is Nothing Then
        For i = 1 To mvarProjects.Count
            updateCount = updateCount + mvarProjects.Item(i).SaveChanges()
        Next i
    Else
        updateCount = -1
    End If
    SaveChanges = updateCount
End Function

Public Sub loadProjectsFromDB(Optional condition As String = " 1=1 ")
    
    Dim projectRecords As Recordset
    Set projectRecords = selectFromDB(dbProjectTable, condition)
    If querySuccess Then
        Set mvarProjects = Nothing
        Set mvarProjects = New Collection
        Dim tmpProject As New Project
        While Not projectRecords.EOF
            tmpProject.initProjectWithRecord projectRecords
            If mvarIsDebug Then tmpProject.debugProject "Displaying Projects"
            mvarProjects.Add tmpProject
            projectRecords.MoveNext
        Wend
        mvarProjectCount = mvarProjects.Count
        mvarIsLoaded = True
    Else
        mvarIsLoaded = False
    End If
End Sub

Public Sub removeProject(ProjectId As Integer)
    Dim tmpProject As New Project
    tmpProject.LoadSingleton (ProjectId)
    If tmpProject.IsLoaded Then
        tmpProject.Delete
        tmpProject.SaveChanges
    End If
End Sub

Public Sub updateProjectInfo(updatedProject As Project)
    updatedProject.SaveChanges
End Sub

Public Function addProject(newProject As Project) As Integer
    If Not newProject Is Nothing Then
       addProject = newProject.SaveChanges
    Else
        addProject = -1
    End If
End Function

Public Function getProjectById(ByVal ProjectId As Integer) As Project
    Dim tmpProject As New Project
    tmpProject.LoadSingleton ProjectId
    Set getProjectById = tmpProject
    
End Function

Public Function getProjects() As Collection
    If Not mvarIsLoaded Then
        loadProjectsFromDB
    End If
    Set getProjects = mvarProjects
    
End Function


Public Property Get Projects() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Projects
    If Not mvarIsLoaded Then
        loadProjectsFromDB
    End If
    Set Projects = mvarProjects
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
    Count = mvarProjectCount
End Property




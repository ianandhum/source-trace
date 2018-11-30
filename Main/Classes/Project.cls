VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarProjectId As Integer 'local copy
Private mvarProjectName As String 'local copy
Private mvarDescription As String 'local copy
Private mvarTags As String 'local copy
Private mvarMode As String 'local copy
Private mvarLocation As String 'local copy
Private mvarCFClass As String 'local copy
Private mvarRemoteId As Integer 'local copy
Private mvarTaskId As Integer 'local copy
Private mvarGitAssoc As Integer 'local copy
Private mvarLastUpdate As Date 'local copy
Private mvarState As String 'local copy

Private mvarIsLoaded As Boolean 'local copy
Private mvarIsDirty As Boolean 'local copy
Private mvarIsNew As Boolean 'local copy
Private mvarIsDeleted As Boolean 'local copy

Private arrayData() As String
Private arrayCols() As String
Private arrayDataTypes() As Integer


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


Public Sub initProjectWithRecord(projectRecord As Recordset)
    Dim errStatus As Boolean
    errStatus = False
    If projectRecord.State = adStateOpen And Not projectRecord.EOF Then
        
        If Not IsNull(projectRecord!project_id) Then
            mvarProjectId = projectRecord!project_id
        Else
            errStatus = True
        End If
        
        If Not IsNull(projectRecord!Name) Then
            mvarProjectName = projectRecord!Name
        Else
            errStatus = True
        End If
        
        If Not IsNull(projectRecord!Description) Then
            mvarDescription = projectRecord!Description
        End If
        
        If Not IsNull(projectRecord!Tags) Then
            mvarTags = projectRecord!Tags
        End If
        
        If Not IsNull(projectRecord!Mode) Then
            mvarMode = projectRecord!Mode
        Else
            errStatus = True
        End If
        
        If Not IsNull(projectRecord!Location) Then
            mvarLocation = projectRecord!Location
        End If
        
        If Not IsNull(projectRecord!cf_class) Then
            mvarCFClass = projectRecord!cf_class
        Else
            errStatus = True
        End If
        
        If Not IsNull(projectRecord!remote_id) Then
            mvarRemoteId = projectRecord!remote_id
        End If
        
        If Not IsNull(projectRecord!task_id) Then
            mvarTaskId = projectRecord!task_id
        End If
        
        If Not IsNull(projectRecord!git_assoc) Then
            mvarGitAssoc = projectRecord!git_assoc
        End If
        
        If Not IsNull(projectRecord!last_update) Then
            mvarLastUpdate = projectRecord!last_update
        Else
            errStatus = True
        End If
        
        If Not IsNull(projectRecord!State) Then
            mvarState = projectRecord!State
        Else
            errStatus = True
        End If
    Else
        errStatus = True
    End If
    mvarIsLoaded = Not errStatus
End Sub

Public Sub debugProject(Optional currProjectName As String = "Project")
    Debug.Print "-----------------------"
    Debug.Print "Debugging Project :: " & currProjectName
    Debug.Print "ProjectId :" & mvarProjectId
    Debug.Print "ProjectName :" & mvarProjectName
    Debug.Print "ProjectDesc :" & mvarDescription
    Debug.Print "ProjectLastUpdate :" & mvarLastUpdate
    Debug.Print "ProjectGitAssoc :" & mvarGitAssoc
    Debug.Print "ProjectClass :" & mvarCFClass
    Debug.Print "ProjectTags :" & mvarTags
    Debug.Print "ProjectLocation :" & mvarLocation
    Debug.Print "ProjectRemoteId :" & mvarRemoteId
    Debug.Print "ProjectTaskId :" & mvarTaskId
    Debug.Print "ProjectMode :" & mvarMode
    Debug.Print "ProjectStatus :" & mvarState
    
    Debug.Print "ProjectIsLoaded :" & mvarIsLoaded
    Debug.Print "ProjectIsDirty :" & mvarIsDirty
    Debug.Print "ProjectIsNew :" & mvarIsNew
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
        SaveChanges = updateDB(arrayCols, arrayData, arrayDataTypes, dbProjectTable, "project_id = " & mvarProjectId)
    ElseIf mvarIsDirty And mvarIsNew Then
        SaveChanges = insertToDB(arrayCols, arrayData, arrayDataTypes, dbProjectTable)
    ElseIf mvarIsDeleted And mvarIsLoaded Then
        SaveChanges = deleteFromDB(dbProjectTable, "project_id = " & mvarProjectId)
    Else
        SaveChanges = -1
    End If
End Function


'Convert Variables to array
Private Sub SerializeForDB()
    arrayData(0) = mvarProjectName
    arrayData(1) = mvarDescription
    arrayData(2) = getNullIfEmpty(mvarTags)
    arrayData(3) = getNullIfEmpty(mvarMode)
    arrayData(4) = getNullIfEmpty(mvarLocation)
    arrayData(5) = getNullIfEmpty(mvarCFClass)
    arrayData(6) = getNullIfEmpty(mvarRemoteId)
    arrayData(7) = getNullIfEmpty(mvarTaskId)
    arrayData(8) = getNullIfEmpty(mvarGitAssoc)
    arrayData(9) = mvarLastUpdate
    arrayData(10) = mvarState
End Sub

Private Sub setDataDefinitions()

    ReDim arrayCols(11)
    arrayCols(0) = "name"
    arrayCols(1) = "description"
    arrayCols(2) = "tags"
    arrayCols(3) = "mode"
    arrayCols(4) = "location"
    arrayCols(5) = "cf_class"
    arrayCols(6) = "remote_id"
    arrayCols(7) = "task_id"
    arrayCols(8) = "git_assoc"
    arrayCols(9) = "last_update"
    arrayCols(10) = "state"

    ReDim arrayDataTypes(11)
    arrayDataTypes(0) = adVarChar
    arrayDataTypes(1) = adVarChar
    arrayDataTypes(2) = adVarChar
    arrayDataTypes(3) = adVarChar
    arrayDataTypes(4) = adVarChar
    arrayDataTypes(5) = adVarChar
    arrayDataTypes(6) = adInteger
    arrayDataTypes(7) = adInteger
    arrayDataTypes(8) = adInteger
    arrayDataTypes(9) = adDate
    arrayDataTypes(10) = adVarChar
    
    ReDim arrayData(11)
    
End Sub


Private Function getNullIfEmpty(data As Variant)
    Dim result As String
    
    If IsEmpty(data) Or data = "" Then
        result = "NULL"
    Else
        result = data
    End If
    getNullIfEmpty = result
End Function

Public Function LoadSingleton(ProjectId As Integer, Optional loadAssocFields As Boolean = True) As Boolean
    If DbConnection.State = adStateOpen Then
        Dim tmpStore As Recordset
        Set tmpStore = selectFromDB(dbProjectTable, " project_id = " & ProjectId)
        If querySuccess Then initProjectWithRecord tmpStore
    End If
End Function

Public Sub GetProgress()
    
End Sub


Public Property Get IsDirty() As Boolean
    IsDirty = mvarIsDirty
End Property


Public Property Get IsLoaded() As Boolean
    IsLoaded = mvarIsLoaded
End Property



Public Property Let State(ByVal vData As String)
    mvarState = vData
    mvarIsDirty = True
End Property


Public Property Get State() As String
    Status = mvarState
End Property



Public Property Let RemoteId(ByVal vData As Integer)
    mvarRemoteId = vData
    mvarIsDirty = True
End Property


Public Property Get RemoteId() As Integer
    RemoteId = mvarRemoteId
End Property

Public Property Let TaskId(ByVal vData As Integer)
    mvarTaskId = vData
    mvarIsDirty = True
End Property


Public Property Get TaskId() As Integer
    TaskId = mvarTaskId
End Property

Public Property Let GitAssoc(ByVal vData As Integer)
    mvarGitAssoc = vData
    mvarIsDirty = True
End Property


Public Property Get GitAssoc() As Integer
    GitAssoc = mvarGitAssoc
End Property


Public Property Let Mode(ByVal vData As String)
    mvarMode = vData
    mvarIsDirty = True
End Property


Public Property Get Mode() As String
    Weight = mvarMode
End Property



Public Property Let Tags(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tags = 5
    mvarTags = vData
    mvarIsDirty = True
End Property


Public Property Get Tags() As String
    Tags = mvarTags
End Property

Public Property Let Location(ByVal vData As String)
    mvarLocation = vData
    mvarIsDirty = True
End Property


Public Property Get Location() As String
    Tags = mvarLocation
End Property



Public Property Let CFClass(ByVal vData As String)
    mvarCFClass = vData
    mvarIsDirty = True
End Property


Public Property Get CFClass() As String
    CFClass = mvarCFClass
End Property


Public Property Let LastUpdate(ByVal vData As Date)
    mvarLastUpdate = vData
    mvarIsDirty = True
End Property


Public Property Get LastUpdate() As Date
    LastUpdate = mvarLastUpdate
End Property



Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
    mvarIsDirty = True
End Property


Public Property Get Description() As String
    Description = mvarDescription
End Property



Public Property Let ProjectName(ByVal vData As String)
    mvarProjectName = vData
    mvarIsDirty = True
End Property


Public Property Get ProjectName() As String
    ProjectName = mvarProjectName
End Property


Public Property Get ProjectId() As Integer
    ProjectId = mvarTaskId
End Property


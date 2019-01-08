VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Store Snippet Information"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSnippetName As String 'local copy
Private mvarDescription As String 'local copy
Private mvarSnippetType As String 'local copy
Private mvarSnippetId As Integer 'local copy
Private mvarLocation As String 'local copy
Private mvarCFClass As String 'local copy
'local variable(s) to hold property value(s)
Private mvarIsNew As Boolean 'local copy
Private mvarIsLoaded As Boolean 'local copy
Private mvarIsDirty As Boolean 'local copy
Private mvarIsDeleted As Boolean 'local copy

Private arrayData() As String
Private arrayCols() As String
Private arrayDataTypes() As Integer


Public Property Let IsDeleted(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsDeleted = 5
    mvarIsDeleted = vData
End Property


Public Property Get IsDeleted() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsDeleted
    IsDeleted = mvarIsDeleted
End Property



Public Property Let IsDirty(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsDirty = 5
    mvarIsDirty = vData
End Property


Public Property Get IsDirty() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsDirty
    IsDirty = mvarIsDirty
End Property



Public Property Let IsLoaded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsLoaded = 5
    mvarIsLoaded = vData
End Property


Public Property Get IsLoaded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsLoaded
    IsLoaded = mvarIsLoaded
End Property



Public Property Let IsNew(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsNew = 5
    mvarIsNew = vData
End Property


Public Property Get IsNew() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsNew
    IsNew = mvarIsNew
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



Public Property Let Location(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
    mvarIsDirty = True
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Let SnippetId(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnippetId = 5
    mvarSnippetId = vData
    mvarIsDirty = True
End Property


Public Property Get SnippetId() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnippetId
    SnippetId = mvarSnippetId
End Property



Public Property Let SnippetType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnippetType = 5
    mvarSnippetType = vData
    mvarIsDirty = True
End Property


Public Property Get SnippetType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnippetType
    SnippetType = mvarSnippetType
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



Public Property Let SnippetName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SnippetName = 5
    mvarSnippetName = vData
    mvarIsDirty = True
End Property


Public Property Get SnippetName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SnippetName
    SnippetName = mvarSnippetName
End Property





Private Sub Class_Initialize()
    setDataDefinitions
End Sub


Public Sub initSnippetWithRecord(SnippetRecord As Recordset)
    Dim errStatus As Boolean
    errStatus = False
    If SnippetRecord.State = adStateOpen And Not SnippetRecord.EOF Then
        If Not IsNull(SnippetRecord!Snippet_id) Then
            mvarSnippetId = SnippetRecord!Snippet_id
        Else
            errStatus = True
        End If
        
        If Not IsNull(SnippetRecord!name) Then
            mvarSnippetName = SnippetRecord!name
        Else
            errStatus = True
        End If
        
        If Not IsNull(SnippetRecord!Description) Then
            mvarDescription = SnippetRecord!Description
        End If
        
        
        If Not IsNull(SnippetRecord!cf_class) Then
            mvarCFClass = SnippetRecord!cf_class
        End If
        
        If Not IsNull(SnippetRecord!Location) Then
            mvarLocation = SnippetRecord!Location
        End If
        
        If Not IsNull(SnippetRecord!Type) Then
            mvarSnippetType = SnippetRecord!Type
        End If
    End If
    mvarIsLoaded = Not errStatus
End Sub

Public Sub debugSnippet(Optional currSnippetName As String = "Snippet")
    Debug.Print "-----------------------"
    Debug.Print "Debugging Snippet :: " & currSnippetName
    Debug.Print "SnippetId :" & mvarSnippetId
    Debug.Print "SnippetName :" & mvarSnippetName
    Debug.Print "SnippetDesc :" & mvarDescription
    Debug.Print "SnippetType :" & mvarSnippetType
    Debug.Print "SnippetClass :" & mvarCFClass
    Debug.Print "SnippetLocation :" & mvarLocation
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
        SaveChanges = updateDB(arrayCols, arrayData, arrayDataTypes, dbSnippetTable, "snippet_id = " & mvarSnippetId)
    ElseIf mvarIsDirty And mvarIsNew Then
        SaveChanges = insertToDB(arrayCols, arrayData, arrayDataTypes, dbSnippetTable)
        Dim rec As New Recordset
        Set rec = runQuery("SELECT MAX(snippet_id) as last_id from tb_snippet")
        If Not rec.EOF Then
            mvarSnippetId = rec!last_id
        End If
        
    ElseIf mvarIsDeleted And mvarIsLoaded Then
        SaveChanges = deleteFromDB(dbSnippetTable, "snippet_id = " & mvarSnippetId)
    Else
        SaveChanges = -1
    End If
End Function


'Convert Variables to array
Private Sub SerializeForDB()
    arrayData(0) = mvarSnippetName
    arrayData(1) = mvarDescription
    arrayData(2) = mvarSnippetType
    arrayData(3) = mvarLocation
    arrayData(4) = getNullIfEmpty(mvarCFClass)
End Sub

Private Sub setDataDefinitions()

    ReDim arrayCols(5)
    arrayCols(0) = "name"
    arrayCols(1) = "description"
    arrayCols(2) = "type"
    arrayCols(3) = "location"
    arrayCols(4) = "cf_class"
    ReDim arrayDataTypes(5)
    arrayDataTypes(0) = adVarChar
    arrayDataTypes(1) = adVarChar
    arrayDataTypes(2) = adVarChar
    arrayDataTypes(3) = adVarChar
    arrayDataTypes(4) = adVarChar
    
    ReDim arrayData(5)
    
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

Public Function LoadSingleton(SnippetId As Integer, Optional loadAssocFields As Boolean = True) As Boolean
    If DbConnection.State = adStateOpen Then
        Dim tmpStore As Recordset
        Set tmpStore = selectFromDB(" [SourceTrace].[dbo].[tb_Snippet] ", " snippet_id = " & SnippetId)
        If querySuccess Then initSnippetWithRecord tmpStore
    End If
End Function


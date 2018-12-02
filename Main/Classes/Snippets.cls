VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snippets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Snippet"
Attribute VB_Ext_KEY = "Member0" ,"Snippet"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection

'local variable(s) to hold property value(s)
Private mvarIsLoaded As Boolean 'local copy
Private mvarSnippetCount As Integer 'local copy
Private mvarIsDebug As Integer 'local copy
Dim tmpSnippet As Snippet


Public Function Add(SnippetName As String, Description As String, SnippetType As String, Location As String, CFClass As String, Optional sKey As String) As Snippet
    'create a new object
    Dim objNewMember As Snippet
    Set objNewMember = New Snippet


    'set the properties passed into the method
    objNewMember.SnippetName = SnippetName
    objNewMember.Description = Description
    objNewMember.SnippetType = SnippetType
    objNewMember.Location = Location
    objNewMember.CFClass = CFClass
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Snippet
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    mvarSnippetCount = 0
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Function SaveChanges() As Integer
    Dim updateCount As Integer
    updateCount = 0
    If mvarIsLoaded And Not mCol Is Nothing Then
        For i = 1 To mCol.Count
            updateCount = updateCount + mCol.Item(i).SaveChanges()
        Next i
    Else
        updateCount = -1
    End If
    SaveChanges = updateCount
End Function

Public Sub loadSnippetsFromDB(Optional condition As String = " 1=1 ")
    
    Dim SnippetRecords As Recordset
    Set SnippetRecords = selectFromDB(dbSnippetTable, condition)
    If querySuccess Then
        While Not SnippetRecords.EOF
            Set tmpSnippet = New Snippet
            tmpSnippet.initSnippetWithRecord SnippetRecords
            If mvarIsDebug Then tmpSnippet.debugSnippet "Displaying Snippets"
            mCol.Add tmpSnippet
            SnippetRecords.MoveNext
            
        Wend
        mvarSnippetCount = mCol.Count
        mvarIsLoaded = True
    Else
        mvarIsLoaded = False
    End If
End Sub

Public Sub removeSnippet(SnippetId As Integer)
    Dim tmpSnippet As New Snippet
    tmpSnippet.LoadSingleton (SnippetId)
    If tmpSnippet.IsLoaded Then
        tmpSnippet.Delete
        tmpSnippet.SaveChanges
    End If
End Sub


Public Sub debugSnippets()
    For i = 1 To mCol.Count
        Debug.Print "-------------"
        Debug.Print "SnippetName: " & mCol.Item(i).SnippetName
        Debug.Print "SnippetId: " & mCol.Item(i).SnippetId
    Next i
    
End Sub

Public Sub updateSnippetInfo(updatedSnippet As Snippet)
    updatedSnippet.SaveChanges
End Sub

Public Function addSnippet(newSnippet As Snippet) As Integer
    If Not newSnippet Is Nothing Then
       addSnippet = newSnippet.SaveChanges
    Else
        addSnippet = -1
    End If
End Function

Public Function getSnippetById(ByVal SnippetId As Integer) As Snippet
    Dim tmpSnippet As New Snippet
    tmpSnippet.LoadSingleton SnippetId
    Set getSnippetById = tmpSnippet
    
End Function

Public Function getSnippets() As Collection
    If Not mvarIsLoaded Then
        loadSnippetsFromDB
    End If
    Set getSnippets = mCol
    
End Function


Public Property Get Snippets() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Snippets
    If Not mvarIsLoaded Then
        loadSnippetsFromDB
    End If
    Set Snippets = mCol
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

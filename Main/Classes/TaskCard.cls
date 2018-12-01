VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCardName As String 'local copy
Private mvarCardDBName As String 'local copy
Private mvarDBIndex As Integer 'local copy
Private mvarCounts As Integer 'local copy
Private mvarIsLoaded As Boolean 'local copy

Public Sub initTaskCard(taskRecord As Recordset)
    Dim errStatus As Boolean
    errStatus = False
    If taskRecord.State = adStateOpen And Not taskRecord.EOF Then
        
        If Not IsNull(taskRecord!card_name) Then
            mvarCardName = taskRecord!card_name
            mvarCardDBName = taskRecord!card_name
        Else
            errStatus = True
        End If
        If Not IsNull(taskRecord!count_nos) Then
            mvarCounts = taskRecord!count_nos
        Else
            errStatus = True
        End If
        mvarIsLoaded = Not errStatus
    End If
End Sub


Public Property Let Counts(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
    mvarCounts = vData
End Property


Public Property Get Counts() As Integer
'used when assigning an Object to the property, on the left side of a Set statement.
    Counts = mvarCounts
End Property


Public Property Let DBIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DBIndex = 5
    mvarDBIndex = vData
End Property

Public Property Get DBIndex() As Integer
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DBIndex = 5
    DBIndex = mvarDBIndex
End Property


Public Property Let CardDBName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CardDBName = 5
    mvarCardDBName = vData
End Property


Public Property Get CardDBName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CardDBName
    CardDBName = mvarCardDBName
End Property



Public Property Let CardName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CardName = 5
    mvarCardName = vData
End Property


Public Property Get CardName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CardName
    CardName = mvarCardName
End Property




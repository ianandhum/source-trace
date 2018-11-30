VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim tvn As Node
Sub RecursePath(ByVal fPath As String, tvw As TreeView)
   Dim File_Name As String
   Dim File_Read As Integer
   Dim strTempPath As String
   Dim i As Integer

   If Right$(fPath, 1) <> "\" Then fPath = fPath & "\"
   
   File_Name = Dir$(fPath, vbDirectory)
   File_Read = 1

   Do While File_Name <> ""
       If File_Name <> "." And File_Name <> ".." Then
           strTempPath = fPath & File_Name
           
           If GetAttr(strTempPath) And vbDirectory Then
               Set tvn = tvw.Nodes.Add(fPath, tvwChild, strTempPath + "\", File_Name)
               tvn.Expanded = True
               RecursePath strTempPath, tvw
           
               File_Name = Dir$(fPath, vbDirectory)
               For i = 2 To File_Read
                   File_Name = Dir$
               Next
           Else
               Set tvn = tvw.Nodes.Add(fPath, tvwChild, strTempPath, File_Name)
           End If
       End If
       
       File_Name = Dir$
       File_Read = File_Read + 1
   Loop
End Sub

Function createTreeView(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer, tvw As TreeView)
    Dim tvn As Node
    Dim tempPath As String
    Dim tempPaths() As String
    
    If Right(path, 1) <> "\" Then path = path & "\"
    
    tempPath = Left$(path, Len(path) - 1)
    tempPaths = Split(tempPath, "\")
    tempPath = tempPaths(UBound(tempPaths))
    Set tvn = tvw.Nodes.Add(, , path, tempPath)
    tvn.Expanded = True
    
    RecursePath path, tvw
    
    
End Function



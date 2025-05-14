VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COMConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"COMConnection"
Attribute VB_Ext_KEY = "Member0" ,"COMConnection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(key As String) As COMConnection
    On Error GoTo AddERR
    
    Set Add = Nothing
    
    'create a new object
    Dim objNewMember As COMConnection
    Set objNewMember = New COMConnection

    objNewMember.key = key
    objNewMember.Index = mCol.Count
       
    mCol.Add objNewMember, key

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    Exit Function
AddERR:
End Function

Public Property Get item(vntIndexKey As Variant) As COMConnection
Attribute item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Error Resume Next
  Set item = mCol(vntIndexKey)
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


'Public Property Get NewEnum() As IUnknown
'    'this property allows you to enumerate
'    'this collection with the For...Each syntax
'    Set NewEnum = mCol.[_NewEnum]
'End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Dim i As Integer
    Dim COMcnt As COMConnection
    
    
    For i = mCol.Count To 1 Step -1
        On Error Resume Next
        mCol.item(i).CloseCOM
        mCol.Remove i
    Next

    Set mCol = Nothing
End Sub


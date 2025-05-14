VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tracciati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private pTracciati  As Collection


Public Function Add(key As String, fieldName As String, fieldDescription As String _
    , fieldType As FIELD_TYPE_ENUM, fieldSizeLen As Long, fieldSizePrecision As Long _
    , fieldFormat As String, fieldRule As String, mappingTable As String, mappingField As String, mappingFieldExternal As String)
    Dim pTracciatoMapping As TracciatoMapping
    Dim pColTracciatoMapping As Collection
    
    Dim pColField As Collection
    Dim pColFieldExt As Collection
    
    On Error GoTo AddERR
    If Not isTracciatoContained(key) Then
        Set pColTracciatoMapping = New Collection
        'creo un collezione sui campi e l'altra sul mapping
        Set pColField = New Collection
        Set pColFieldExt = New Collection
        
        pColTracciatoMapping.Add pColField, "FIELD"
        pColTracciatoMapping.Add pColFieldExt, "FIELDEXT"
        
        pTracciati.Add pColTracciatoMapping, key
    Else
        Set pColTracciatoMapping = pTracciati(key)
        
        Set pColField = pColTracciatoMapping("FIELD")
        Set pColFieldExt = pColTracciatoMapping("FIELDEXT")

    End If
    
    
    Set pTracciatoMapping = New TracciatoMapping
    With pTracciatoMapping
        .fieldName = "Col" & fieldName
        .fieldDescription = fieldDescription
        .fieldType = fieldType
        .fieldSizeLen = fieldSizeLen
        .fieldSizePrecision = fieldSizePrecision
        .fieldFormat = fieldFormat
        .fieldRule = fieldRule
        .mappingTable = mappingTable
        mappingField = Split(mappingField, ",")(0)
        .mappingField = mappingField
        .mappingFieldExternal = mappingFieldExternal
    End With
    
    pColFieldExt.Add pTracciatoMapping, mappingFieldExternal
    pColField.Add pTracciatoMapping, mappingField
    
    Set Add = pTracciatoMapping
    Exit Function
AddERR:
End Function

Public Property Get Item(vntIndexKey As Variant) As Collection
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  
  Set Item = pTracciati(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = pTracciati.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    pTracciati.Remove vntIndexKey
End Sub


'Public Property Get NewEnum() As IUnknown
'    'this property allows you to enumerate
'    'this collection with the For...Each syntax
'    Set NewEnum = mCol.[_NewEnum]
'End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set pTracciati = New Collection
End Sub

Public Function isTracciatoContained(Tracciato As String) As Boolean
    Dim pColTracciatoMapping As Collection
    
    On Error Resume Next
    
    Set pColTracciatoMapping = Nothing
    Set pColTracciatoMapping = pTracciati(Tracciato)
    
    
   isTracciatoContained = Not pColTracciatoMapping Is Nothing
End Function


Public Function isContained(Tracciato As String, fieldName As String) As Boolean
    Dim pColTracciatoMapping As Collection
    Dim pTracciatoMapping As TracciatoMapping
    Dim pColField As Collection
    Dim pColFieldExt As Collection
    
    On Error Resume Next
    
    Set pTracciatoMapping = Nothing
    Set pColTracciatoMapping = pTracciati(Tracciato)
    
    Set pColField = pColTracciatoMapping("FIELD")
    Set pColFieldExt = pColTracciatoMapping("FIELDEXT")
    
    Set pTracciatoMapping = pColField(fieldName)
    If Err.Number > 0 Then
        Set pTracciatoMapping = pColFieldExt(fieldName)
    End If
    
    
   isContained = Not pTracciatoMapping Is Nothing
End Function




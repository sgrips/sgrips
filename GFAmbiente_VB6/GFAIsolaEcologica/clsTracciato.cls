VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TracciatoMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private pFieldName As String
Private pFieldDescription As String
Private pFieldType As FIELD_TYPE_ENUM
Private pFieldSizeLen As Long
Private pFieldSizePrecision As Long
Private pFieldFormat As String
Private pFieldRule As String

Private pMappingTable As String
Private pMappingField As String
Private pMappingFieldExternal As String


Public Property Get fieldName() As String
    fieldName = pFieldName
End Property

Public Property Let fieldName(Value As String)
    pFieldName = Value
End Property

Public Property Get fieldDescription() As String
    fieldDescription = pFieldDescription
End Property

Public Property Let fieldDescription(Value As String)
    pFieldDescription = Value
End Property

Public Property Get fieldType() As FIELD_TYPE_ENUM
    fieldType = pFieldType
End Property

Public Property Let fieldType(Value As FIELD_TYPE_ENUM)
    pFieldType = Value
End Property

Public Property Get fieldSizeLen() As Long
    fieldSizeLen = pFieldSizeLen
End Property

Public Property Let fieldSizeLen(Value As Long)
    pFieldSizeLen = Value
End Property

Public Property Get fieldSizePrecision() As Long
    fieldSizePrecision = pFieldSizePrecision
End Property

Public Property Let fieldSizePrecision(Value As Long)
    pFieldSizePrecision = Value
End Property


Public Property Get fieldFormat() As String
    fieldFormat = pFieldFormat
End Property

Public Property Let fieldFormat(Value As String)
    pFieldFormat = Value
End Property

Public Property Get fieldRule() As String
    fieldRule = pFieldRule
End Property

Public Property Let fieldRule(Value As String)
    pFieldRule = Value
End Property

Public Property Get mappingTable() As String
    mappingTable = pMappingTable
End Property

Public Property Let mappingTable(Value As String)
    pMappingTable = Value
End Property

Public Property Get mappingField() As String
    mappingField = pMappingField
End Property

Public Property Let mappingField(Value As String)
    pMappingField = Value
End Property

Public Property Get mappingFieldExternal() As String
    mappingFieldExternal = pMappingFieldExternal
End Property

Public Property Let mappingFieldExternal(Value As String)
    pMappingFieldExternal = Value
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSONValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarAttributo As String 'local copy
Private mvarValore As Variant 'local copy
Public Property Let Valore(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Valore = 5
    mvarValore = vData
End Property


Public Property Set Valore(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Valore = Form1
    Set mvarValore = vData
End Property


Public Property Get Valore() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Valore
    If IsObject(mvarValore) Then
        Set Valore = mvarValore
    Else
        Valore = mvarValore
    End If
End Property



Public Property Let Attributo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Attributo = 5
    mvarAttributo = vData
End Property


Public Property Get Attributo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Attributo
    Attributo = mvarAttributo
End Property





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private valori As New ColJSONValue

Public Function ToString() As String
    Dim item As JSONValue
    Dim strJson As String
    
    For Each item In valori
        If TypeOf item.Valore Is json Then
            Dim jvalue As json
            Set jvalue = item.Valore
            
            strJson = Chr$(34) & item.Attributo & Chr$(34) & ":" & jvalue.ToString()
        Else
            strJson = strJson & Chr$(34) & item.Attributo & Chr$(34) & ":" & Chr$(34) & item.Valore & Chr$(34) & ","
        End If
    Next
    
    If (Right(strJson, 1) = ",") Then strJson = left(strJson, Len(strJson) - 1)
    
    ToString = "{" & strJson & "}"
End Function


Public Sub Add(ByVal nomeAttributo As String, valoreAttributo As Variant)

    If valori.item(nomeAttributo) Is Nothing Then
        valori.Add nomeAttributo, valoreAttributo
    Else
        valori.item(nomeAttributo).Valore = valoreAttributo
    End If
End Sub


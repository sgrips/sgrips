VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsWord"
Attribute VB_Ext_KEY = "Member0" ,"clsWord"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variabile locale per memorizzare l'insieme.
Private mCol As Collection
'Per generare questo evento, utilizzare l'istruzione RaiseEvent con la seguente sintassi:
'RaiseEvent DocQuit[(arg1, arg2, ... , argn)]
Public Event DocQuit()


Public Function Add(Document As Word.Application, sKey As String) As clsword
    'crea un nuovo oggetto
    Dim objNewMember As clsword
    Set objNewMember = New clsword

    Set objNewMember.Document = Document
    objNewMember.Key = sKey
    Set objNewMember.Parent = Me
    
    mCol.Add objNewMember, sKey
    
    'restituisce l'oggetto creato
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsword
Attribute Item.VB_UserMemId = 0
    'Utilizzato per fare riferimento a un elemento nell'insieme.
    'vntIndexKey contiene la chiave o l'indice dell'insieme,
    'e per questo motivo è dichiarata come Variant.
    'Sintassi: Set foo = x.Item(xyz) oppure Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'Utilizzata per recuperare il numero di elementi dell'insieme.
    'Sintassi: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'Utilizzata per rimuovere un elemento dall'insieme.
    'vntIndexKey contiene l'indice o la chiave, e per questo
    'motivo viene dichiarata come Variant.
    'Sintassi: x.Remove(xyz)
    mCol.Remove vntIndexKey
    If Me.Count = 0 Then RaiseEvent DocQuit
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Questa proprietà consente di enumerare l'insieme
    'corrente con la sintassi For...Each.
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Crea l'insieme quando viene creata questa classe.
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'Rimuove l'insieme quando la classe viene eliminata.
    Set mCol = Nothing
End Sub


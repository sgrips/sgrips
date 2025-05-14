VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColImagesRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ImmaginiRD"
Attribute VB_Ext_KEY = "Member0" ,"ImmaginiRD"
Option Explicit

'Variabile locale per memorizzare l'insieme.
Private mCol As Collection

Public Function Add(nameImage As String, sKey As String, Des As String) As ImagesRD
    'crea un nuovo oggetto
    Dim objNewMember As ImagesRD
    
        
    
    Set objNewMember = New ImagesRD
    'imposta le proprietà passate al metodo
    objNewMember.nameImage = nameImage
    objNewMember.key = sKey
    objNewMember.Description = Des
    
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set objNewMember.ColImagesRD = Me

    'restituisce l'oggetto creato
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Function Contains(vntIndexKey As String) As Boolean
    On Error Resume Next
    Dim obj As Object
    Contains = False
    Set obj = Item(vntIndexKey)
    If Err.Number = 0 Then Contains = True
    
End Function

Public Property Get Item(vntIndexKey As Variant) As ImagesRD
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
End Sub


'Public Property Get NewEnum() As IUnknown
    'Questa proprietà consente di enumerare l'insieme
    'corrente con la sintassi For...Each.
'    Set NewEnum = mCol.[_NewEnum]
'End Property


Private Sub Class_Initialize()
    'Crea l'insieme quando viene creata questa classe.
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'Rimuove l'insieme quando la classe viene eliminata.
    Set mCol = Nothing
End Sub


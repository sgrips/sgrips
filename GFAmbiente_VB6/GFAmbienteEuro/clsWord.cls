VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private WithEvents mvarDocument As Word.Application  'Copia locale.
Attribute mvarDocument.VB_VarHelpID = -1
Private mvarKey As String
'Per generare questo evento, utilizzare l'istruzione RaiseEvent con la seguente sintassi:
'RaiseEvent DocQuit[(arg1, arg2, ... , argn)]
Public Event DocQuit()
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarParent As clsWords 'Copia locale.
Friend Property Set Parent(ByVal vData As clsWords)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.Parent = Form1
    Set mvarParent = vData
End Property


Friend Property Get Parent() As clsWords
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Parent
    Set Parent = mvarParent
End Property


Public Property Set Document(ByVal vData As Word.Application)
'Utilizzato per l'assegnazione di un oggetto a una proprietà, a sinistra di un'istruzione Set.
'Syntax: Set x.Document = Form1
    Set mvarDocument = vData
End Property


Public Property Get Document() As Word.Application
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Document
    Set Document = mvarDocument
End Property


Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let Key(ByVal s As String)
    mvarKey = s
End Property

Private Sub Class_Initialize()

End Sub

Private Sub mvarDocument_Quit()
    Parent.Remove Me.Key
End Sub

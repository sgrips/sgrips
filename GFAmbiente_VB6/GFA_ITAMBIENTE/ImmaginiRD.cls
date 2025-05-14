VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagesRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Immagine raccolta differenziata"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColImagesRD"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarnameImage As String 'Copia locale.
Private mvarkey As String 'Copia locale.
Private mvarColImagesRD As ColImagesRD
Private mvarLinkOBJ As CheckBox
Private mvarDes As String


Public Property Get ColImagesRD() As ColImagesRD
    If mvarColImagesRD Is Nothing Then
        Set mvarColImagesRD = New ColImagesRD
    End If


    Set ColImagesRD = mvarColImagesRD
End Property


Public Property Set ColImagesRD(vData As ColImagesRD)
    Set mvarColImagesRD = vData
End Property
Private Sub Class_Terminate()
    Set mvarColImagesRD = Nothing
End Sub


Public Property Set LinkOBJ(ByRef vData As CheckBox)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.key = 5
    Set mvarLinkOBJ = vData
End Property


Public Property Get LinkOBJ() As CheckBox
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.key
    Set LinkOBJ = mvarLinkOBJ
End Property


Public Property Let key(ByVal vData As String)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.key = 5
    mvarkey = vData
End Property


Public Property Get Description() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.key
    Description = mvarDes
End Property

Public Property Let Description(ByVal vData As String)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.key = 5
    mvarDes = vData
End Property


Public Property Get key() As String
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.key
    key = mvarkey
End Property




Public Property Let nameImage(ByVal vData As String)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.nameImage = 5
    mvarnameImage = vData
End Property


Public Property Get nameImage() As String
Attribute nameImage.VB_UserMemId = 0
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.nameImage
    nameImage = mvarnameImage
End Property




VERSION 5.00
Begin VB.UserControl UCPremi 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14640
   ScaleHeight     =   1605
   ScaleWidth      =   14640
   Begin VB.TextBox txtPuntiPerPremio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10680
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtDescrizione 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   8175
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lblImage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click per Modifica"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12600
      MouseIcon       =   "UCPremi.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Image PicImage 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   12600
      MouseIcon       =   "UCPremi.ctx":0152
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "UCPremi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dichiarazioni di eventi:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Viene generato non appena un form viene visualizzato o quando le dimensioni di un oggetto vengono modificate."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Viene generato quando si preme e quindi si rilascia un pulsante del mouse su un oggetto."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Viene generato quando si preme e si rilascia due volte in rapida successione un pulsante del mouse su un oggetto."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Viene generato quando si preme un tasto mentre lo stato attivo si trova su un oggetto."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Viene generato quando si preme e si rilascia un tasto ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Viene generato quando si rilascia un tasto mentre lo stato attivo si trova su un oggetto."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Viene generato quando si preme il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Viene generato quando si sposta il mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Viene generato quando si rilascia il pulsante del mouse mentre lo stato attivo si trova su un oggetto."




Private Sub lblImage_Click()
    Dim s As FileSystemObject
    Dim sFile As String
    DoCommonDialog DT_Open, sFile, "DDD", "Selezionare la foto del premio", UserControl.hwnd, "Premi(*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp", 1, App.Path, "*.jpg"
    
    
    If sFile <> vbNullString Then
        Set s = New FileSystemObject
        If Not s.FolderExists(App.Path & "\Premi") Then
            s.CreateFolder App.Path & "\Premi"
        End If
        s.CopyFile sFile, App.Path & "\Premi\" & s.GetFileName(sFile)
        PicImage.Picture = LoadPicture(App.Path & "\Premi\" & s.GetFileName(sFile))
        PicImage.ToolTipText = s.GetFileName(sFile)
        PropertyChanged "Foto"
    End If
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Restituisce o imposta il colore di sfondo utilizzato per la visualizzazione di testo e grafica in un oggetto."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Restituisce o imposta il colore di primo piano utilizzato per la visualizzazione di testo e grafica in un oggetto."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Restituisce un oggetto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indica se il controllo Label o lo sfondo di un controllo Shape è trasparente oppure opaco."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Restituisce o imposta lo stile del bordo di un oggetto."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Ridisegna completamente un oggetto."
    UserControl.Refresh
End Sub



Private Sub PicImage_Click()
    lblImage_Click
End Sub

Private Sub txtCode_Change()
    PropertyChanged "CodePremio"
End Sub

Private Sub txtDescrizione_Change()
    PropertyChanged "Descrizione"
End Sub

Private Sub txtPuntiPerPremio_Change()
    PropertyChanged "PremioDa"
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    txtCode.Text = PropBag.ReadProperty("CodePremio", "")
    PicImage.ToolTipText = PropBag.ReadProperty("Foto", "")
    txtPuntiPerPremio.Text = PropBag.ReadProperty("PremioDa", 1)
    txtDescrizione.Text = PropBag.ReadProperty("Descrizione", "")
End Sub

Private Sub UserControl_Resize()
    PicImage.Left = UserControl.ScaleWidth - PicImage.Width - 150
    txtPuntiPerPremio.Left = PicImage.Left - txtPuntiPerPremio.Width - 150
    txtDescrizione.Width = txtPuntiPerPremio.Left - 150 - txtDescrizione.Left
    
    RaiseEvent Resize

End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("CodePremio", txtCode.Text, "")
    Call PropBag.WriteProperty("Foto", PicImage.ToolTipText, "")
    Call PropBag.WriteProperty("PremioDa", txtPuntiPerPremio.Text, 0)
    Call PropBag.WriteProperty("Descrizione", txtDescrizione.Text, "")
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=txtDescrizione,txtDescrizione,-1,Text
Public Property Get Descrizione() As String
Attribute Descrizione.VB_Description = "breve descrizione del prodotto"
Attribute Descrizione.VB_MemberFlags = "14"
    Descrizione = txtDescrizione.Text
End Property

Public Property Let Descrizione(ByVal New_Descrizione As String)
    txtDescrizione.Text = New_Descrizione
    PropertyChanged "Descrizione"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=txtCode,txtCode,-1,Text
Public Property Get CodePremio() As String
Attribute CodePremio.VB_Description = "nome del premio"
Attribute CodePremio.VB_MemberFlags = "14"
    CodePremio = txtCode.Text
End Property

Public Property Let CodePremio(ByVal New_CodePremio As String)
    txtCode.Text = New_CodePremio
    PropertyChanged "CodePremio"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=PicImage,PicImage,-1,ToolTipText
Public Property Get Foto() As String
Attribute Foto.VB_Description = "Foto del premio"
Attribute Foto.VB_UserMemId = 0
Attribute Foto.VB_MemberFlags = "14"
    Foto = PicImage.ToolTipText
End Property

Public Property Let Foto(ByVal New_Foto As String)
    PicImage.ToolTipText = New_Foto
    If New_Foto <> "" Then
        Set PicImage.Picture = LoadPicture(App.Path & "\Premi\" & New_Foto)
    End If
    PropertyChanged "Foto"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=SliderDa,SliderDa,-1,Value
Public Property Get PremioDa() As Long
Attribute PremioDa.VB_MemberFlags = "14"
    PremioDa = txtPuntiPerPremio.Text
End Property

Public Property Let PremioDa(ByVal New_PremioDa As Long)
    txtPuntiPerPremio.Text = New_PremioDa
    PropertyChanged "PremioDa"
End Property

VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmHTTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WebService - ITAmbiente"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   4920
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   8415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
   Begin InetCtlsObjects.Inet HTTP 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
End
Attribute VB_Name = "frmHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim operation As Integer
Dim pURL As String
Dim pURLToken As String
Dim pEsitoInvio As EnumStatoInvio

'accesso
Dim pIstat_comune As String
Dim pId_movim As String
Dim pCod_pda As String
Dim pFlag_ospite As String
Dim pCfisc_operaz  As String
Dim pDat_movim As String
Dim pOra_movim As String
Dim pBarcode_cfisc As String

Property Let WS_Istat_comune(s As String)
     pIstat_comune = s
End Property
Property Let WS_Id_movim(s As String)
     pId_movim = s
End Property
Property Let WS_Cod_pda(s As String)
     pCod_pda = s
End Property
Property Let WS_Flag_ospite(s As String)
     pFlag_ospite = s
End Property
Property Let WS_CodFisc(s As String)
     pCfisc_operaz = s
End Property
Property Let WS_Data_movim(s As String)
     pDat_movim = s
End Property
Property Let WS_Ora_movim(s As String)
     pOra_movim = s
End Property
Property Let WS_Barcode_cfisc(s As String)
     pBarcode_cfisc = s
End Property

Property Let URLToken(s As String)
     pURLToken = s
End Property


Property Let URLWebService(s As String)
     pURL = s
End Property


Private Sub cmdOK_Click()
 cmdOK.enabled = False
 Unload Me
End Sub


Public Function DoWork(ope As Integer) As EnumStatoInvio
    operation = ope
    Timer1.enabled = True
   
    If operation = 2 Then
        Load Me
        cmdOK.enabled = False
        Do
            DoEvents
        Loop While Not cmdOK.enabled
        
    Else
        Me.Show vbModal
    End If
    
    DoWork = pEsitoInvio
End Function


Private Sub Connect()
    pEsitoInvio = Invio_Errore_Generico
    
    Status "Connessione al server:" & pURL
    DoEvents

    HTTP.URL = pURL
    HTTP.Protocol = IIf(InStr(pURL, "https://") <> -1, icHTTPS, icHTTP)
    
    On Error GoTo ConnectERR
    Exit Sub
ConnectERR:
    Status "HTTP Errore:" & Err.Description
End Sub

Private Sub SendFile()
    On Error GoTo SendFileERR
    
    Dim sHeader As String
    Dim sPostData As String
    Dim json As New json
    Dim json1 As New json
    
    Connect
    cmdOK.enabled = False
    txtInfo = ""
    
    sHeader = "Content-Type: application/json" & vbCrLf & "User-Agent: Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)" & vbCrLf & "access_token:3b08b906be9188efe5f962119269b39e" & vbCrLf
        
    'invio HTTP
    json1.Add "istat_comune", pIstat_comune
    json1.Add "id_movim", pId_movim
    json1.Add "cod_pda", pCod_pda
    json1.Add "flag_ospite", pFlag_ospite
    json1.Add "cfisc_operaz", pCfisc_operaz
    json1.Add "dat_movim", pDat_movim
    json1.Add "ora_movim", pOra_movim
    json1.Add "Barcode_cfisc", pBarcode_cfisc
    
    json.Add "pda", json1
    sPostData = json.ToString()
    
    Status "Invio JSON : " & sPostData
    
    HTTP.Execute , "POST", sPostData, sHeader
    
    Do While HTTP.StillExecuting
        DoEvents
    Loop
    
    cmdOK.enabled = True
    If pEsitoInvio = Invio_Ok Then cmdOK_Click
    
    Exit Sub
SendFileERR:
    pEsitoInvio = Invio_Errore_Generico
    cmdOK.enabled = True
    Status "Errore HTTP:" & Err.Description
End Sub

Private Sub Status(s As String)
    txtInfo = s & vbCrLf & txtInfo.Text
End Sub


Private Sub HTTP_StateChanged(ByVal state As Integer)
 ' Recupera la risposta del server utilizzando
   ' il metodo GetChunk quando State = 12.

   Dim vtData As Variant ' Variabile Data.
   vtData = ""
   Select Case state
   ' ... Gli altri casi sono omessi.
   Case icError ' 11
      ' In caso di errore, restituisce ResponseCode e
      ' ResponseInfo.
      vtData = "Stato:Errore " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
      pEsitoInvio = Invio_Errore_Generico
       
   Case StateConstants.icConnected
    vtData = "Stato:Connesso " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icConnecting
    vtData = "Stato:In Connessione " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icDisconnected
    vtData = "Stato:Disconnesso " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icDisconnecting
    vtData = "Stato:In Disconnessione " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icHostResolved
    vtData = "Stato:HOST Risolto " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icNone
   Case StateConstants.icRequesting
    vtData = "Stato:In Richiesta " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icRequestSent
    vtData = "Stato:Invio Richiesta " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icResolvingHost
    vtData = "Stato:Ricerca HOST " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icReceivingResponse
    vtData = "Stato:In Ricezione Risposta " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   Case StateConstants.icResponseReceived
    vtData = "Stato:Risposta Ricevuta " & HTTP.ResponseCode & ":" & _
      HTTP.ResponseInfo
   
   Case icResponseCompleted ' 12

      Dim strData As String
      Dim bDone As Boolean: bDone = False

      ' Recupera il primo blocco.
      vtData = HTTP.GetChunk(1024, icString)
      DoEvents

      Do While Not bDone
         strData = strData & vtData
         ' Recupera il blocco successivo.
         vtData = HTTP.GetChunk(1024, icString)
         DoEvents

         If Len(vtData) = 0 Then
            bDone = True
         End If
      Loop
      
      If Len(strData) > 0 Then
         If (InStr(UCase(strData), "OK") > 0) Then
            pEsitoInvio = Invio_Ok
        Else
            pEsitoInvio = Invio_Errore_Rifiutato
        End If
         vtData = "Esito Invio:" & strData
      End If
      vtData = vtData & " ResponseCompleted"
        
      
      Case Else
         vtData = "Stato:" & state & " " & HTTP.ResponseCode & ":" & HTTP.ResponseInfo
      
   
   End Select
   
   Status CStr(vtData)
  
End Sub

Private Sub Timer1_Timer()
    Timer1.enabled = False
    Select Case operation
        Case 1, 2
            SendFile
        Case 3
            
    End Select
End Sub

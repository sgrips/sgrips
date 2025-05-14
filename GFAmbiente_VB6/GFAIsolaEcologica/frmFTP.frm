VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmFTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP- ITAmbiente"
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
   Begin InetCtlsObjects.Inet FTP 
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
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim operation As Integer
Dim pFilesource As String, pFileDestination As String, pRename As Boolean
Dim pListafile() As String
Dim pEsitoInvio As Boolean
Dim pErrore As Boolean

Property Let FileSource(s As String)
     pFilesource = s
End Property

Property Let FileDestination(s As String)
     pFileDestination = s
End Property

Property Get FileDestination() As String
       FileDestination = pFileDestination
End Property


Property Let Rename(s As Boolean)
     pRename = s
End Property


Private Sub cmdOK_Click()
 cmdOK.enabled = False
 Unload Me
End Sub


Public Function DoWork(ope As Integer) As Boolean
    pErrore = False
    operation = ope
    Timer1.enabled = True
    Me.Show vbModal
    DoWork = pEsitoInvio
End Function



Private Sub Connect()
    Status "Connessione al server:" & gFTP_URL
    
    DoEvents

    FTP.URL = gFTP_URL
    FTP.UserName = gFTP_UID
    FTP.password = gFTP_PWD
    On Error GoTo ConnectERR
    
    Exit Sub
ConnectERR:
    Status "FTP Errore:" & Err.Description
End Sub

Private Sub SendFile()
    On Error GoTo SendFileERR
    Dim tmpFile As String
    
    Connect
     cmdOK.enabled = False
    txtInfo = ""
    'FTP.Execute , "SEND C:\Sviluppi\SviluppiStudio60\GFA\Esportazione\leggi.txt out/leggi.txt"
    Dim fo As FileSystemObject
    Set fo = New FileSystemObject
    
    If pRename Then
        tmpFile = Replace(pFileDestination, fo.GetExtensionName(pFileDestination), "tmp")
    Else
        tmpFile = pFileDestination
    End If
    
        
    Status "Invio file :" & pFilesource & ", " & tmpFile
    
    FTP.Execute , "SEND " & pFilesource & " " & tmpFile
    pEsitoInvio = False
    
    Do While FTP.StillExecuting
        DoEvents
    Loop
    
    If pRename Then
    
        cmdOK.enabled = False
        Status "Rinomina File:" & tmpFile & " in " & pFileDestination
        FTP.Execute , "RENAME " & tmpFile & " " & pFileDestination
        
        Do While FTP.StillExecuting
            DoEvents
        Loop
        pEsitoInvio = True
        
    End If
    
    cmdOK.enabled = True
     
    If Not pErrore Then cmdOK_Click
     
    Exit Sub
SendFileERR:
     cmdOK.enabled = True
     
     Status "Errore FTP:" & Err.Description
End Sub

Private Sub Status(s As String)
    txtInfo = s & vbCrLf & txtInfo.Text
End Sub

Private Sub FTP_StateChanged(ByVal state As Integer)
 ' Recupera la risposta del server utilizzando
   ' il metodo GetChunk quando State = 12.

   Dim vtData As Variant ' Variabile Data.
   vtData = ""
   Select Case state
   ' ... Gli altri casi sono omessi.
   
   
   
   Case icError ' 11
      ' In caso di errore, restituisce ResponseCode e
      ' ResponseInfo.
      vtData = "Stato:Errore " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
      pErrore = True
   Case StateConstants.icConnected
    vtData = "Stato:Connesso " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icConnecting
    vtData = "Stato:In Connessione " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icDisconnected
    vtData = "Stato:Disconnesso " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icDisconnecting
    vtData = "Stato:In Disconnessione " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icHostResolved
    vtData = "Stato:HOST Risolto " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icNone
   Case StateConstants.icRequesting
    vtData = "Stato:In Richiesta " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icRequestSent
    vtData = "Stato:Invio Richiesta " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icResolvingHost
    vtData = "Stato:Ricerca HOST " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icReceivingResponse
    vtData = "Stato:In Ricezione Risposta " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   Case StateConstants.icResponseReceived
    vtData = "Stato:Risposta Ricevuta " & FTP.ResponseCode & ":" & _
      FTP.ResponseInfo
   
   Case icResponseCompleted ' 12

      Dim strData As String
      Dim bDone As Boolean: bDone = False

      ' Recupera il primo blocco.
      vtData = FTP.GetChunk(1024, icString)
      DoEvents

      Do While Not bDone
         strData = strData & vtData
         ' Recupera il blocco successivo.
         vtData = FTP.GetChunk(1024, icString)
         DoEvents

         If Len(vtData) = 0 Then
            bDone = True
            
         End If
      Loop
      
      If Len(strData) > 0 Then
         pListafile = Split(strData, Chr(13) + Chr(10))
         ReDim Preserve pListafile(UBound(pListafile) - 1)
         vtData = "Trovato file:" & Join(pListafile, ",")
      End If
       
      vtData = vtData & " ResponseCompleted"
              
      pErrore = False
      
   Case Else
      vtData = "Stato:" & state & " " & FTP.ResponseCode & ":" & FTP.ResponseInfo
   End Select
   
   Status CStr(vtData)
  
End Sub

Private Sub DownloadFile()
    Dim pathIn As String
    Dim so As Scripting.FileSystemObject
    
    pathIn = "in/"
    
    cmdOK.enabled = False
    On Error GoTo DownloadFileErr
    Connect
    
    FTP.Execute , "DIR " & pFilesource
    
    Do While FTP.StillExecuting
        DoEvents
    Loop
    
    Set so = New FileSystemObject
    On Error Resume Next
    Dim numfile As Integer
    numfile = 0
    numfile = UBound(pListafile)
    On Error GoTo DownloadFileErr
    If numfile > 0 Then
        Status "Ricezione file :" & pathIn & pListafile(0) & " in " & pFileDestination & pListafile(0)
     
        FTP.Execute , "GET " & pathIn & pListafile(0) & " " & pFileDestination & pListafile(0)
     
     
        Do While FTP.StillExecuting
            DoEvents
        Loop
        
        If pRename Then
        
            cmdOK.enabled = False
            Status "Rinomina File:" & pathIn & pListafile(0) & " in " & pathIn & Replace(pListafile(0), so.GetExtensionName(pListafile(0)), "") & Format(Now, "YYYYMMDD_HHNNSS")
 
            FTP.Execute , "RENAME " & pathIn & pListafile(0) & " " & pathIn & Replace(pListafile(0), so.GetExtensionName(pListafile(0)), "") & Format(Now, "YYYYMMDD_HHNNSS")
            
            Do While FTP.StillExecuting
                DoEvents
            Loop
        End If
            
        FileDestination = pFileDestination & pListafile(0)
    Else
       FileDestination = ""
    End If
    
    cmdOK.enabled = True
 
    If Not pErrore Then cmdOK_Click
    
    Exit Sub
    
DownloadFileErr:
FileDestination = ""
     cmdOK.enabled = True
        Status "Errore FTP:" & Err.Description
End Sub

Private Sub Timer1_Timer()
    Timer1.enabled = False
    
    Select Case operation
        Case 1
            SendFile
        Case 2
            DownloadFile
    End Select
End Sub

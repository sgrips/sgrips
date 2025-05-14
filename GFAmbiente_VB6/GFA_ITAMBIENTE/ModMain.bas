Attribute VB_Name = "ModMain"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Enum ComandiBilanciaEnum
    cmd_RichiestaPesa
    cmd_ScritturaAlibi
    cmd_CambioCanale
    cmd_Tara
End Enum

Public Enum EnumStatoInvio
    Invio_Errore_Generico = 0
    Invio_Errore_Rifiutato = 1
    Invio_Ok = 2
End Enum

Public Enum ProtocolloInvioEnum
    Protocollo_Invio_Nessuo = 1
    Protocollo_Invio_HTTP = 2
    Protocollo_Invio_FTP = 3
End Enum

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" _
  (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function getParent Lib "user32" _
  Alias "GetParent" (ByVal hWnd As Long) As Long

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public mDEBUG As Boolean


Public Const WS_CAPTION = &HC00000 'Window with a title bar and border
Public Const WS_THICKFRAME = &H40000 'Create a window with a sizing border.
Public Const WS_SYSMENU = &H80000 'Create a window with a system menu on its
'Title bar. Must be combined with WS_CAPTION
Public Const MY_WS_DOCK = WS_CAPTION Or WS_THICKFRAME Or WS_SYSMENU
Public Const GWL_STYLE = (-16)

Public hOldParent As Long

Public Declare Function SetWindowPos Lib "user32" _
  (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
  ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOZORDER = &H4



Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public colTracciati As Tracciati


Public Enum FIELD_TYPE_ENUM
    FIELD_TYPE_CHAR = 0
    FIELD_TYPE_NUM = 1
    FIELD_TYPE_BOOL = 2
End Enum

Public g_rsPese As ADODB.Recordset


'messaggio pesa
Public PesaMSGTipoOpeDa   As Long, _
        PesaMSGTipoOpeNrCaratteri  As Long, _
        PesaMSGDataDa  As Long, PesaMSGDataNrCaratteri  As Long, _
        PesaMSGOraDa  As Long, PesaMSGOraNrCaratteri  As Long, _
        PesaMSGProgressivoDa  As Long, PesaMSGProgressivoNrCaratteri   As Long, _
        PesaMSGCodRifiutoDa  As Long, PesaMSGCodRifiutoNrCaratteri   As Long, _
        PesaMSGCodBadgeDa   As Long, PesaMSGCodBadgeNrCaratteri  As Long, _
        PesaMSGPesoIngressoDa1 As Long, PesaMSGPesoIngressoDa1NrCaratteri1 As Long, _
        PesaMSGPesoIngressoDa2 As Long, PesaMSGPesoIngressoDa2NrCaratteri2 As Long, _
        PesaMSGPesoUscitaDa1 As Long, PesaMSGPesoUscitaDa1NrCaratteri1 As Long, _
        PesaMSGPesoUscitaDa2 As Long, PesaMSGPesoUscitaDa2NrCaratteri2 As Long, _
        PesaMSGPesoNetto As Long, PesaMSGPesoNettoNrCaratteri As Long

Public PesaMSGTipoOpeDa_2   As Long, _
        PesaMSGTipoOpeNrCaratteri_2  As Long, _
        PesaMSGDataDa_2  As Long, PesaMSGDataNrCaratteri_2  As Long, _
        PesaMSGOraDa_2  As Long, PesaMSGOraNrCaratteri_2  As Long, _
        PesaMSGProgressivoDa_2  As Long, PesaMSGProgressivoNrCaratteri_2   As Long, _
        PesaMSGCodRifiutoDa_2  As Long, PesaMSGCodRifiutoNrCaratteri_2   As Long, _
        PesaMSGCodBadgeDa_2   As Long, PesaMSGCodBadgeNrCaratteri_2  As Long, _
        PesaMSGPesoIngressoDa1_2 As Long, PesaMSGPesoIngressoDa1NrCaratteri1_2 As Long, _
        PesaMSGPesoIngressoDa2_2 As Long, PesaMSGPesoIngressoDa2NrCaratteri2_2 As Long, _
        PesaMSGPesoUscitaDa1_2 As Long, PesaMSGPesoUscitaDa1NrCaratteri1_2 As Long, _
        PesaMSGPesoUscitaDa2_2 As Long, PesaMSGPesoUscitaDa2NrCaratteri2_2 As Long, _
        PesaMSGPesoNetto_2 As Long, PesaMSGPesoNettoNrCaratteri_2 As Long


Public PesaMSGEcoDataDa As Long, _
        PesaMSGEcoDataNrCaratteri As Long, _
        PesaMSGEcoOraDa As Long, _
        PesaMSGEcoOraNrCaratteri As Long, _
        PesaMSGEcoCodRifiutoDa  As Long, _
        PesaMSGEcoCodRifiutoNrCaratteri As Long, _
        PesaMSGEcoCodBadgeDa As Long, _
        PesaMSGEcoCodBadgeNrCaratteri As Long, _
        PesaMSGEcoPesoNettoDa1 As Long, _
        PesaMSGEcoPesoNettoNrCaratteri1 As Long

Public PesaMSGLAumas400PesoNettoDa As Long, _
       PesaMSGLAumas400PesoNettoCaratteri As Long, _
       PesaMSGLAumas400PesoNettoDaDec As Long, _
       PesaMSGLAumas400PesoNettoDecCaratteri As Long

Public PesaMSGRETELPesoNettoDa As Long, _
       PesaMSGRETELPesoNettoCaratteri As Long, _
       PesaMSGRETELPesoNettoDaDec As Long, _
       PesaMSGRETELPesoNettoDecCaratteri As Long, _
       PesaMSGRETELPesoStabileDa As Long, _
       PesaMSGRETELPesoStabileCaratteri As Long


'parametri di input
Public gStartMode As Long
Public gTimerLettura As Long
Public gCODISTAT As String

'percorso file di lavoro
Public gPathWrkFile As String

Private mGGOldFile As Long
Private mFileNumber As Long
Private mCOMconnections As New COMConnections
Private mCOM As COMConnection

Public pfrmGrafici(2) As frmGrafici
Public pfrmSocieta As FrmSocietaTab
Public pfrmScuole As FrmScuoleTab
Public pfrmPremi As frmPremi
Public pfrmMovimentiScuole As frmMovimentiScuole


Public pfrmDestinatari As frmDestinatari
Public pfrmTrasportatori As frmTrasportatori
Public pfrmUtenti As frmUtenti
Public pfrmUser As frmUser
Public pfrmCodiciCER As frmCodiciCER
Public pfrmPesata As frmPesata
Public pfrmProtezioneDB As frmProtezioneDB
Public pfrmGiacenze As frmGiacenze
Public pfrmTotMovimentiBonusPunti As frmTotMovimentati
Public pfrmRegistri As frmRegistri
Public pfrmUscita  As frmUscita
Public pfrmFormulari As frmFormulari
Public pfrmIntermediari As frmIntermediari
Public pfrmImpImmobiliUtenti As frmImpImmobiliUtenti

Public gIdImpiantoProduttore As Long
Public gIdSocietaProduttore As Long


Public gRegCaricoSuRegistroOgni As Integer
Public gRegCaricoSuRegistroAllUscita As Integer
Public gQtaMaxGiornaliera As Boolean
Public gGGAllarmeGiacNormativa As Integer
Public gGGAllarmeGiacGestionale As Integer

Public gFattureFile As String
Public gFattureFileType As Integer
Public gFattureSeparatore As String

Public gMovimentiFile As String
Public gMovimentiFileType As Integer
Public gMovimentiSeparatore As String

Public gCodiciCERFile As String
Public gCodiciCERFileType As Integer
Public gCodiciCERSeparatore As String

Public gUtentiFile As String
Public gUtentiFileType As Integer
Public gUtentiSeparatore As String

Public gAccessiFile As String
Public gAccessiFileType As Integer
Public gAccessiSeparatore As String


Public gAnagrafiche As Boolean


Public gPercorsoImp As String
Public gPercorsoEsp As String
Public gRichiedereAUscitaProgramma As Boolean
Public gPercorsoBackup As String
Public gPercorsoStoricizzazione As String
Public gStoricizzazioneAutomatica As Boolean
Public gReportFormulari As String
Public gFTP_URL As String
Public gFTP_UID As String
Public gFTP_PWD As String
Public gFTP_Active As Boolean

'HTTP
Public gTipoInvioAccessi As ProtocolloInvioEnum
Public gURLTokenSicurezza As String
Public gURLWebService As String


Public pImpImmUtePath As String
Public pImpImmUteSepCol As String
Public pImpImmUteSepDec As String
Public pImpImmUteSepText As String


Public gStampaPesataAnteprima As Single
Public gStampaPesataStampante  As String
Public gStampaPesataNrCopie  As Integer
Public gStampaPesataMargSx As Double
Public gStampaPesataMargDx As Double
Public gStampaPesataMargUp As Double
Public gStampaPesataMargDwn As Double

Public gStampaRegistroAnteprima As Single
Public gStampaRegistroStampante  As String
Public gStampaRegistroNrCopie  As Integer
Public gStampaRegistroMargSx As Double
Public gStampaRegistroMargDx As Double
Public gStampaRegistroMargUp As Double
Public gStampaRegistroMargDwn As Double


Public gStampaFormularioAnteprima As Single
Public gStampaFormularioStampante  As String
Public gStampaFormularioNrCopie  As Integer
Public gStampaFormularioMargSx As Double
Public gStampaFormularioMargDx As Double
Public gStampaFormularioMargUp As Double
Public gStampaFormularioMargDwn As Double

Public gStampaScontrinoAnteprima As Single
Public gStampaScontrinoStampante  As String
Public gStampaScontrinoNrCopie  As Integer
Public gStampaScontrinoMargSx As Double
Public gStampaScontrinoMargDx As Double
Public gStampaScontrinoMargUp As Double
Public gStampaScontrinoMargDwn As Double
Public gStampaScontrinoInsDichiaraDiResponsabilita As Single
Public gStampaScontrinoInsComuneImpiantoScontrino As Single
Public gStampaScontrinoInsImpiantoScontrino As Single
Public gStampaScontrinoInsRagSocScontrino As Single
Public gStampaScontrinoDichiarazione As String
Public gLunghezzaBadge As Long


Public gCodCausaleMovManuale As String
Public gCodImmobileMovManuale As Long
Public gCodUserMovManuale As String
Public gCodUtenteMovManuale  As Long
Public gPesaIsolaCod As String
Public gCodCausaleMovEcoPuntoIN As String
Public gCodCausaleMovEcoPuntoOUT As String
Public gflgMovECOPuntoGeneraGiacenza As Integer
Public gflgCodPesaECOPunto As Integer
Public gflgEcoPuntoDataMovimentoIns As Integer
Public gflgImmobili As Integer
Public gflgOrderByImmobiliUD As Integer



Public gflgFatturazioneConKD As Integer
Public gDisabilitaNotifiche As Integer
Public gflgAttiviMUD As Boolean
Public gPuntiAssegnatiPer As Long
    

'accessi
Public gActiveAccess As Integer
Public gIdentifyUserByCodFiscalePiva As Integer
Public gMaxUsersIsland As Integer
Public gRecycleTimeUsersIsland As Integer
Public gAccessDeviceCOMName As String
Public gSecRefreshAccessMonitor As Integer


'costanti
Private Const PATTERNFILELOG = "L???????.ASC"

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SW_SHOWNA = 8
Public Const SW_HIDE = &H0
Public Const SWP_SHOWWINDOW = &H40
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNORMAL = 1
Public Const LEN_ID_PARCEL = 5
Public Const SW_MINIMIZE = 6
Public Const STARTF_USESHOWWINDOW = &H1





Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (lpsz1 As Any, lpsz2 As Any) As Long



'Messaggio Pesa
Type messaggioPesa
    TipoOperazione As String * 2
    Data As Date
    Progressivo As Integer
    CodRifiuto As String * 10
    DesRifiuto As String * 20
    CodCliente  As String * 10
    DesCliente  As String * 20
    CodGenOperazione As String * 10
    DesGenOperazione As String * 20
    CodAct As String * 16
    DesAct As String * 20
    codBadge As String * 16
    Campo1 As String * 22
    Campo2 As String * 22
    Campo3 As String * 22
    Campo4 As String * 22
    Campo5 As String * 22
    Campo6 As String * 22
    Campo7 As String * 22
    Risultato As String * 10
    NumeroColli As Integer
    PesoLordo As Double
    UMPesoLordo As String * 2
    Tara  As Double
    UMTara As String * 2
    PesoNetto As Double
    UMPesoNetto As String * 2
    PesoStabile As Boolean
End Type


'tipi entità

Public Enum enumEntita
    FATTURE
    MOVIMENTI
    FORMULARIO
    UTENTI
    REGISTRO
    IMMOBILI
    DOCUMENTO_PESATA_IN
    DOCUMENTO_PESATA_OUT
End Enum

Public Const NUMERATORE_IMMOBILE_MIN = 59999

'tipi log
Enum TIPO_LOG
    TPL_ERRORE = 0
    TPL_INFORAMITIVO = 1
End Enum

'declarations for working with Ini filesPrivate
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

'Database
Public gDBC  As ADODB.Connection
Public gJRO As JRO.JetEngine
Public gPathDB  As String
Public gPasswordDB  As String


Public Type currentPesa
    ModelloPesa As String
    ModalitaLetturaPesoAutomatica As Boolean
    BadgeInStringaPesata As Boolean
    MSGPerDoppiaPesa As Boolean
    
    CmdSceltaPesa As String
    CmdSceltaPesaRetOKCode As String
    cmdLetturaPeso As String
    CmdScritturaAlibiM As String
    CmdScritturaAlibiMRetERRCode As String
    cmdPesoStabileCode As String
    CmdTara As String
    
End Type

Private mCurrentPesa As currentPesa


Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function


Public Function GetComuneFromCode(codComune As String) As Recordset
    Dim rs As ADODB.Recordset

    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    rs.Source = "select * from comuni where codIstat='" & codComune & "'"
    rs.Open
    Set rs.ActiveConnection = Nothing
    Set GetComuneFromCode = rs.Clone
    rs.Close

    
End Function

Public Function GetCurrentComune() As Recordset
    Dim rs As ADODB.Recordset
    Dim codComune As String
    
    If gCODISTAT = vbNullString Then
        Set rs = gDBC.Execute("select codComune from societa where idSocieta=" & gIdSocietaProduttore & " And codSoggetto='S'")
        If rs.EOF Then MsgBox "Attenzione occorre inserire prima la società": Exit Function
        codComune = rs(0)
        rs.Close
        
    Else
        codComune = gCODISTAT
    End If

    Set GetCurrentComune = GetComuneFromCode(codComune)
    
End Function



Public Sub letturaStrutturaMessaggioPesa()
    Dim l As Long
    Dim s As String
    Dim sp() As String
    
    
    'doppia pesa
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "TipoOperazione", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGTipoOpeDa_2 = sp(0)
    PesaMSGTipoOpeNrCaratteri_2 = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "Data", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGDataDa_2 = sp(0)
    PesaMSGDataNrCaratteri_2 = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "Ora", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGOraDa_2 = sp(0)
    PesaMSGOraNrCaratteri_2 = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "Progressivo", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGProgressivoDa_2 = sp(0)
    PesaMSGProgressivoNrCaratteri_2 = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "CodRifiuto", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGCodRifiutoDa_2 = sp(0)
    PesaMSGCodRifiutoNrCaratteri_2 = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "CodBadge", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGCodBadgeDa_2 = sp(0)
    PesaMSGCodBadgeNrCaratteri_2 = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "PesoIngresso1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa1_2 = sp(0)
    PesaMSGPesoIngressoDa1NrCaratteri1_2 = sp(1)


    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "PesoIngresso2", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa2_2 = sp(0)
    PesaMSGPesoIngressoDa2NrCaratteri2_2 = sp(1)


    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "PesoUscita1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa1_2 = sp(0)
    PesaMSGPesoIngressoDa1NrCaratteri1_2 = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "PesoUscita2", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa2_2 = sp(0)
    PesaMSGPesoIngressoDa2NrCaratteri2_2 = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaDoppiaPesata", "PesoNetto", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoNetto_2 = sp(0)
    PesaMSGPesoNettoNrCaratteri_2 = sp(1)
        
    'una pesa
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "TipoOperazione", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGTipoOpeDa = sp(0)
    PesaMSGTipoOpeNrCaratteri = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "Data", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGDataDa = sp(0)
    PesaMSGDataNrCaratteri = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "Ora", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGOraDa = sp(0)
    PesaMSGOraNrCaratteri = sp(1)
    
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "Progressivo", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGProgressivoDa = sp(0)
    PesaMSGProgressivoNrCaratteri = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "CodRifiuto", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGCodRifiutoDa = sp(0)
    PesaMSGCodRifiutoNrCaratteri = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "CodBadge", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
     sp = Split(s, ",")
    PesaMSGCodBadgeDa = sp(0)
    PesaMSGCodBadgeNrCaratteri = sp(1)
    
    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "PesoIngresso1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa1 = sp(0)
    PesaMSGPesoIngressoDa1NrCaratteri1 = sp(1)


    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "PesoIngresso2", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa2 = sp(0)
    PesaMSGPesoIngressoDa2NrCaratteri2 = sp(1)


    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "PesoUscita1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa1 = sp(0)
    PesaMSGPesoIngressoDa1NrCaratteri1 = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "PesoUscita2", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoIngressoDa2 = sp(0)
    PesaMSGPesoIngressoDa2NrCaratteri2 = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("Struttura", "PesoNetto", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGPesoNetto = sp(0)
    PesaMSGPesoNettoNrCaratteri = sp(1)
    
    
    
    
    'Struttura ECOPUNTO
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaECO", "Data", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGEcoDataDa = sp(0)
    PesaMSGEcoDataNrCaratteri = sp(1)


    s = Space$(255)
    l = GetPrivateProfileString("StrutturaECO", "Ora", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGEcoOraDa = sp(0)
    PesaMSGEcoOraNrCaratteri = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaECO", "CodBadge", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGEcoCodBadgeDa = sp(0)
    PesaMSGEcoCodBadgeNrCaratteri = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaECO", "CodRifiuto", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGEcoCodRifiutoDa = sp(0)
    PesaMSGEcoCodRifiutoNrCaratteri = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaECO", "PesoNetto1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGEcoPesoNettoDa1 = sp(0)
    PesaMSGEcoPesoNettoNrCaratteri1 = sp(1)

    'Struttura Pesa Aumas400
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaAumas400", "PesoNetto1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGLAumas400PesoNettoDa = sp(0)
    PesaMSGLAumas400PesoNettoCaratteri = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaAumas400", "PesoNetto1Dec", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGLAumas400PesoNettoDaDec = sp(0)
    PesaMSGLAumas400PesoNettoDecCaratteri = sp(1)

    'Struttura Pesa RETEL
    s = Space$(255)
    l = GetPrivateProfileString("StrutturaRETEL", "PesoNetto1", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGRETELPesoNettoDa = sp(0)
    PesaMSGRETELPesoNettoCaratteri = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaRETEL", "PesoNetto1Dec", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGRETELPesoNettoDaDec = sp(0)
    PesaMSGRETELPesoNettoDecCaratteri = sp(1)

    s = Space$(255)
    l = GetPrivateProfileString("StrutturaRETEL", "PesoStabile", 0, s, Len(s), App.Path & "\MessaggioPesa.ini")
    s = left$(s, l)
    sp = Split(s, ",")
    PesaMSGRETELPesoStabileDa = sp(0)
    PesaMSGRETELPesoStabileCaratteri = sp(1)



End Sub


Public Sub Main()
    Dim i As Integer
    Dim pSplash  As Form
    Dim P As Printer
    
    'abilita la lettura simulata delle COM
    mDEBUG = False
    
    InitCommonControlsVB
    
    
    
    'command parameters
    'mode=1 normal,2 mobile - user=nome utente,pwd=password utente
    If Command$ <> vbNullString Then
        Dim sParams() As String
        sParams = Split(Command$ & " ", " ")
        
        For i = LBound(sParams) To UBound(sParams)
            Dim sParamValue() As String
            
            If sParams(i) = vbNullString Then Exit For
            sParamValue = Split(sParams(i), "=")
            
            If UCase$(sParamValue(0)) = "MODE" Then
                gStartMode = sParamValue(1)
            ElseIf UCase$(sParamValue(0)) = "USER" Then
                Dim m_codUser As String
                m_codUser = sParamValue(1)
            ElseIf UCase$(sParamValue(0)) = "PWD" Then
                Dim m_PWD As String
                m_PWD = getTextDecrypt(sParamValue(1))
            ElseIf UCase$(sParamValue(0)) = "CODISTAT" Then
                gCODISTAT = sParamValue(1)
            ElseIf UCase$(sParamValue(0)) = "NOCOM" Then
                Dim m_NOCOM As Boolean
                m_NOCOM = True
                MsgBox "Disabilitate COM"
            ElseIf UCase$(sParamValue(0)) = "ANAGRAFICHE" Then
                gAnagrafiche = IIf(sParamValue(1) = 1, True, False)
            End If
        Next
    End If
    
    
    Set pSplash = New frmSplash
    On Error GoTo MainERR
    gPathWrkFile = App.Path & "\Temp"
    
  '  pSplash.Show
      
     
    pSplash.setInfo "Inizializzazione in corso . . ."
    DoEvents
    
    If FindWindow(ByVal 0&, ByVal "frmIsolaEcologicaAttiva") > 0 Then Unload pSplash: MsgBox "Gestione Isola Ecologica già attivo", vbCritical: End
    
    Set gJRO = CreateObject("jro.JetEngine")
    
    Screen.MousePointer = vbHourglass
    
    Dim l As Long
    gReportFormulari = Space(255)
    l = GetPrivateProfileString("Report", "ReportFormulario", "", gReportFormulari, Len(gReportFormulari), App.Path & "\config.ini")
    If (l > 0) Then
        gReportFormulari = left$(gReportFormulari, l)
    Else
        gReportFormulari = "Reportformulario"
    End If
    
    pSplash.setInfo "Apertura Connessione DB . . ."
    
    ApriConnessione
    
    gIdSocietaProduttore = GetSetting(App.Title, "IMPOSTAZIONI", "SOCIETA", -1)
    gIdImpiantoProduttore = GetSetting(App.Title, "IMPOSTAZIONI", "IMPIANTO", -1)
    
    
    MDIMain.Show
    
    If gAnagrafiche Then
        MDIMain.mnuDocumenti.Visible = False
        MDIMain.mnuDestinatari.Visible = False
        MDIMain.mnuTrasportatori.Visible = False
        MDIMain.mnuIntermediari.Visible = False
        MDIMain.mnuPesate.Visible = False
        MDIMain.mnuMovUscita.Visible = False
    End If
    
    MDIMain.mnuFinestra.Visible = False
     
    pSplash.setInfo "Pulizia LOG . . ."
    
    CancelloLogFile
    
    ScriviLog TPL_INFORAMITIVO, App.Title, "Applicazione  Avviata", "Main"
       
       
    SaveSetting App.Title, "IMPOSTAZIONI", "PERCORSOAPPLICAZIONE", App.Path
    gRegCaricoSuRegistroOgni = GetSetting(App.Title, "IMPOSTAZIONI", "RECARICOSUREGISTROOGNI", 2)
    gRegCaricoSuRegistroAllUscita = GetSetting(App.Title, "IMPOSTAZIONI", "RECARICOSUREGISTROALLUSCITA", 0)
    gQtaMaxGiornaliera = ConvToBoolean(GetSetting(App.Title, "IMPOSTAZIONI", "QTAMASSIMAGIORNALIERA", True))


    
    gGGAllarmeGiacNormativa = GetSetting(App.Title, "IMPOSTAZIONI", "GGREMINDERGIACNORMATIVA", 334)
    gGGAllarmeGiacGestionale = GetSetting(App.Title, "IMPOSTAZIONI", "GGREMINDERGIACGESTIONALE", 5)
    
    gFattureFile = GetSetting(App.Title, "IMPOSTAZIONI", "FATTUREFILE", "FATT")
    gFattureFileType = GetSetting(App.Title, "IMPOSTAZIONI", "FATTUREFILETYPE", "0")
    gFattureSeparatore = GetSetting(App.Title, "IMPOSTAZIONI", "FATTUREFILESEPARATORE", "|")

    gMovimentiFile = GetSetting(App.Title, "IMPOSTAZIONI", "MOVIMENTIFILE", "MOV")
    gMovimentiFileType = GetSetting(App.Title, "IMPOSTAZIONI", "MOVIMENTIFILETYPE", "1")
    gMovimentiSeparatore = GetSetting(App.Title, "IMPOSTAZIONI", "MOVIMENTIFILESEPARATORE", ";")

    gCodiciCERFile = GetSetting(App.Title, "IMPOSTAZIONI", "CODICICERFILE", "CER")
    gCodiciCERFileType = GetSetting(App.Title, "IMPOSTAZIONI", "CODICICERFILETYPE", "1")
    gCodiciCERSeparatore = GetSetting(App.Title, "IMPOSTAZIONI", "CODICICERFILESEPARATORE", ";")

    gUtentiFile = GetSetting(App.Title, "IMPOSTAZIONI", "UTENTIFILE", "UTE")
    gUtentiFileType = GetSetting(App.Title, "IMPOSTAZIONI", "UTENTIFILETYPE", "1")
    gUtentiSeparatore = GetSetting(App.Title, "IMPOSTAZIONI", "UTENTIFILESEPARATORE", ";")

    gAccessiFile = GetSetting(App.Title, "IMPOSTAZIONI", "ACCESSIFILE", "ACC")
    gAccessiFileType = GetSetting(App.Title, "IMPOSTAZIONI", "ACCESSIFILETYPE", "1")
    gAccessiSeparatore = GetSetting(App.Title, "IMPOSTAZIONI", "ACCESSIFILESEPARATORE", ";")


    gPercorsoEsp = GetSetting(App.Title, "IMPOSTAZIONI", "PERCORSOESP", App.Path & "\Esportazione")
    gPercorsoImp = GetSetting(App.Title, "IMPOSTAZIONI", "PERCORSOIMP", App.Path & "\Importazione")
    
    gFTP_URL = GetSetting(App.Title, "IMPOSTAZIONI", "FTPURL", "")
    gFTP_UID = GetSetting(App.Title, "IMPOSTAZIONI", "FTPUID", "")
    gFTP_PWD = GetSetting(App.Title, "IMPOSTAZIONI", "FTPPWD", "")
    gFTP_Active = ConvToBoolean(GetSetting(App.Title, "IMPOSTAZIONI", "FTPACTIVE", 0))
    
    'HTTP
    gTipoInvioAccessi = GetSetting(App.Title, "IMPOSTAZIONI", "TIPOINVIOACCESSI", "2")
    gURLTokenSicurezza = GetSetting(App.Title, "IMPOSTAZIONI", "URLTOKENSICUREZZA", "3b08b906be9188efe5f962119269b39e")
    gURLWebService = GetSetting(App.Title, "IMPOSTAZIONI", "URLWEBSERVICE", "https://cornaredo.e-portal.it/api/waste/pda/waste_ccr_access")

    
    gStampaPesataAnteprima = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATAANTEPRIMA", "1")
    gStampaPesataStampante = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATASTAMPANTE", Printer.DeviceName)
    gStampaPesataNrCopie = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATANRCOPIE", "2")

    gStampaRegistroAnteprima = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTROANTEPRIMA", "1")
    gStampaRegistroStampante = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTROSTAMPANTE", Printer.DeviceName)
    gStampaRegistroNrCopie = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTRONRCOPIE", "1")

    gStampaFormularioAnteprima = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOANTEPRIMA", "1")
    gStampaFormularioStampante = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOSTAMPANTE", Printer.DeviceName)
    gStampaFormularioNrCopie = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIONRCOPIE", "1")
    
    
    gStampaPesataMargSx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGSX", 5)
    gStampaPesataMargDx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGDX", 5)
    gStampaPesataMargUp = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGUP", 10)
    gStampaPesataMargDwn = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAPESATAMARGDWN", 10)

    gStampaScontrinoAnteprima = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOANTEPRIMA", "1")
    gStampaScontrinoStampante = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOSTAMPANTE", Printer.DeviceName)
    gStampaScontrinoInsDichiaraDiResponsabilita = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGDICHIARAZIONE", 1)
    
    gStampaScontrinoInsComuneImpiantoScontrino = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGCOMUNEIMPIANTO", 1)
    gStampaScontrinoInsImpiantoScontrino = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGIMPIANTO", 1)
    gStampaScontrinoInsRagSocScontrino = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOFLGRAGSOC", 1)
    
    gStampaScontrinoDichiarazione = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINODICHIARAZIONE", "Dichiaro che i rifiuti conferiti provengono dall'immobile di mia proprietà sito nel Comune di @CODCOMUNE@." & _
                " Dichiaro altresì di aver preso visione delle indicazioni  previste dal" & _
                " regolamento, dei rischi specifici e dei comportamenti di sicurezza" & _
                " del Centro di Raccolta e di averle chiaramente comprese.")
    
    
    gActiveAccess = GetSetting(App.Title, "IMPOSTAZIONI", "ACTIVEACCESS", 0)
    gIdentifyUserByCodFiscalePiva = GetSetting(App.Title, "IMPOSTAZIONI", "IDENTIFYUSERBYCODFISCALEPIVA", 0)
    gMaxUsersIsland = GetSetting(App.Title, "IMPOSTAZIONI", "MAXUSERSISLAND", 30)
    gRecycleTimeUsersIsland = GetSetting(App.Title, "IMPOSTAZIONI", "RECYCLETIMEUSERSISLAND", 30)
    gAccessDeviceCOMName = GetSetting(App.Title, "IMPOSTAZIONI", "ACCESSDEVICECOMNAME", "COM1")
    gSecRefreshAccessMonitor = GetSetting(App.Title, "IMPOSTAZIONI", "SECREFRESHACCESSMONITOR", 5)
    
    gLunghezzaBadge = GetSetting(App.Title, "IMPOSTAZIONI", "LUNGHEZZABADGE", "13")
    
    gStampaScontrinoNrCopie = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINONRCOPIE", 2)
    gStampaScontrinoMargSx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGSX", 0)
    gStampaScontrinoMargDx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGDX", 0)
    gStampaScontrinoMargUp = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGUP", 0)
    gStampaScontrinoMargDwn = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPASCONTRINOMARGDWN", 0)

    gStampaRegistroMargSx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGSX", 5)
    gStampaRegistroMargDx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGDX", 5)
    gStampaRegistroMargUp = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGUP", 10)
    gStampaRegistroMargDwn = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAREGISTROMARGDWN", 10)

    gStampaFormularioMargSx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGSX", 5)
    gStampaFormularioMargDx = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGDX", 5)
    gStampaFormularioMargUp = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGUP", 10)
    gStampaFormularioMargDwn = GetSetting(App.Title, "IMPOSTAZIONI", "STAMPAFORMULARIOMARGDWN", 10)
    
    gCodCausaleMovManuale = GetSetting(App.Title, "IMPOSTAZIONI", "CODCAUSALEMOVDAUSCITAPIU", "99")
    gCodUtenteMovManuale = GetSetting(App.Title, "IMPOSTAZIONI", "CODUTENTEMOVDAUSCITAPIU", 0)
    gCodImmobileMovManuale = GetSetting(App.Title, "IMPOSTAZIONI", "CODIMMOBILEMOVDAUSCITAPIU", 0)
    gCodUserMovManuale = GetSetting(App.Title, "IMPOSTAZIONI", "CODUSERMOVDAUSCITAPIU", 0)
    
    
    
    gCodCausaleMovEcoPuntoIN = GetSetting(App.Title, "IMPOSTAZIONI", "CODCAUSALEMOVECOPUNTOIN", "03")
    gCodCausaleMovEcoPuntoOUT = GetSetting(App.Title, "IMPOSTAZIONI", "CODCAUSALEMOVECOPUNTOOUT", "04")
    gflgMovECOPuntoGeneraGiacenza = GetSetting(App.Title, "IMPOSTAZIONI", "FLGMOVECOPUNTOGENMOV", 0)
    gflgCodPesaECOPunto = GetSetting(App.Title, "IMPOSTAZIONI", "FLGCODPESAECOPUNTO", 0)
    gflgEcoPuntoDataMovimentoIns = GetSetting(App.Title, "IMPOSTAZIONI", "FLGECOPUNTODATAMOVIMENTOINS", 0)
    gflgOrderByImmobiliUD = GetSetting(App.Title, "IMPOSTAZIONI", "FLGORDERBYIMMOBILIUD", 1)
    
    gPesaIsolaCod = GetSetting(App.Title, "IMPOSTAZIONI", "CODPESAISOLA", "00")
     
    gflgImmobili = GetSetting(App.Title, "IMPOSTAZIONI", "GESTIONEIMMOBILI", 1)
    gflgFatturazioneConKD = GetSetting(App.Title, "IMPOSTAZIONI", "FATTURAZIONECONKD", 0)
    
    gDisabilitaNotifiche = GetSetting(App.Title, "IMPOSTAZIONI", "NOTIFICHE", 0)
    gPuntiAssegnatiPer = GetSetting(App.Title, "IMPOSTAZIONI", "PUNTIASSEGNATIPER", 0)
    
    
    gPercorsoBackup = GetSetting(App.Title, "DATABASE", "PERCORSOBACKUP", App.Path & "\Esportazione")
    gPercorsoStoricizzazione = GetSetting(App.Title, "DATABASE", "PERCORSOSTORICO", App.Path & "\Esportazione")
    gRichiedereAUscitaProgramma = ConvToBoolean(GetSetting(App.Title, "DATABASE", "RICHIESTABACKUPUSCITA", True))
    gStoricizzazioneAutomatica = ConvToBoolean(GetSetting(App.Title, "DATABASE", "STORICIZZAZIONEAUTO", True))
    
    Screen.MousePointer = vbDefault
    

    pSplash.setInfo "Generazione Registrazioni di Carico"
    
    If gDBC Is Nothing Then
        pSplash.Visible = False
        MDIMain.apriDataBase
        pSplash.Visible = True
    End If
    
        
    generazioneRegistrazioneDiCarico
    
    pSplash.setInfo "Carica Tracciati"
    
    caricaTracciati
    
    Set g_rsPese = CreateObject("ADODB.Recordset")
    g_rsPese.LockType = adLockBatchOptimistic
    g_rsPese.CursorLocation = adUseClient
    g_rsPese.CursorType = adOpenStatic
    
    Set g_rsPese.ActiveConnection = gDBC
    g_rsPese.Source = "Select * From Pese order by codPesa"
    g_rsPese.Open
    Set g_rsPese.ActiveConnection = Nothing
    g_rsPese.Sort = "codPesa"
    If Command$ <> "NOCOM" Then InitPese True, pSplash
    
    Unload pSplash
    
 
    
    Set pSplash = Nothing
    
   
    If m_codUser <> "" Then
        Dim error_code As Long
        frmLogin.IsValidUser m_codUser, m_PWD, error_code, gCODISTAT
    Else
        frmLogin.ShowME GetSetting(App.Title, "IMPOSTAZIONI", "LASTUSER", "User")
    End If
    
    If frmLogin.LoginSucceded = False Then Err.Raise 9999, "Main", "Utente non valido"
    
    
    If frmLogin.codUser = "GFA" Then MDIMain.attivaMenuSegreti
        
    If GetSetting(App.Title, "IMPOSTAZIONI", "VISUALIZZAALLARMEGIACENZA", True) Then
        DevEnvMDB.cmdGiacenzeDaRitirare gGGAllarmeGiacGestionale, gGGAllarmeGiacNormativa
        
        If (Not DevEnvMDB.rscmdGiacenzeDaRitirare.EOF) Then
            If frmMessageBox.ShowME("Attenzione esistono rifiuti che superano la " & _
                "data di giacenza prevista!!" & vbCrLf & vbCrLf & "Analizza le attuali giacenze?", "VISUALIZZAALLARMEGIACENZA") = vbYes Then frmGiacenze.Show
        End If
    End If
    
    'se l'anno corrente non è inserito tra quelli che devono essere storicizzati lo inserico
    Dim rs As Recordset
    Set rs = gDBC.Execute("select * from storico where anno =" & CStr(Year(Now)))
    
    If rs.EOF Then
        gDBC.Execute ("insert into storico (anno) values(" & CStr(Year(Now)) & ")")
    End If
    
    
    If gStoricizzazioneAutomatica Then
        Set rs = gDBC.Execute("select * from storico where anno =" & CStr(Year(Now) - 1) & " and Storicizzato=false and DateAdd('m',1,[LastAlert]) <#" & Format(Now, "yyyy\-MM\-dd HH\:mm\:ss") & "#")
        If (Not rs.EOF) Then
            If frmMessageBox.ShowME("Attenzione, l'anno contabile " & CStr(Year(Now) - 1) & " è terminato !" & _
                vbCrLf & vbCrLf & "Si vuole procedere con la storicizzazione dei dati?", vbNullString) = vbYes Then
                frmStoricizzazione.Show
            Else
                gDBC.Execute ("update storico set LastAlert =#" & Format(Now, "yyyy\-MM\-dd HH\:mm\:ss") & "# where anno =" & CStr(Year(Now) - 1) & " and Storicizzato=false")
            End If
        End If
    End If

    'verifico se i numeratori devono essere azzerati perchè trascorso un anno contabile
    Set rs = gDBC.Execute("select * from contatori where anno =" & CStr(Year(Now) - 1))
    If (Not rs.EOF) Then
        If MsgBox("Attenzione prima di poter movimentare della nuova merce, occorre azzerare i Numeratori per il nuovo anno contabile." + vbCrLf + _
                  "Questo è possibile farlo confermando ora l'operazione o utilizzando la funzione di menu per la storicizzazione dati." + vbCrLf + _
                  "Si vuole procedere ora con l'azzeramento ?", vbQuestion + vbYesNo) = vbYes Then
            gDBC.Execute ("delete * from Contatori where Entita<>'IMMOBILI' and Entita<>'UTENTI'")
        Else
            Unload MDIMain
            Exit Sub
        End If
        
    End If

    rs.Close
    Set rs = Nothing
    DoEvents
    
    If DevEnvMDB.rscmdSocieta.state = adStateClosed Then
        DevEnvMDB.cmdSocieta
    End If
    
    If DevEnvMDB.rscmdSocieta.EOF Then
        MDIMain.CaricaSocieta
  
    Else
    
        'se primo lancio vanno imposti parametri
        If gIdSocietaProduttore = -1 Or gIdImpiantoProduttore = -1 Or gCodImmobileMovManuale = 0 Then
            MsgBox "Impostare i dati della societa, dell'impianto produttore e l'informazioni per la movimentazione." + vbCrLf + "Impostare la tipologia di pesa collegata all'impianto" + vbCrLf
            frmImpostazioni.Show vbModal
        Else
             If gStartMode = 1 And Forms.Count <= 3 Then MDIMain.CaricaFormPesa
        End If
    End If
    
    MDIMain.tmrInvioAccessi = True
    
    Exit Sub
MainERR:
    If Not pSplash Is Nothing Then Unload pSplash
    DoEvents
    MsgBox "Main:" & Error$, vbCritical + vbOKOnly, App.Title
    If Not MDIMain Is Nothing Then Unload MDIMain
End Sub



Public Function getCurrentPesa(Optional NoChangeRow As Boolean = False) As currentPesa
    Static StaticCodPesa As String
    
    
    If StaticCodPesa <> gPesaIsolaCod Then
        If NoChangeRow = False Then
            g_rsPese.Filter = ""
            g_rsPese.Filter = "codPesa='" & gPesaIsolaCod & "'"
        End If
        
        mCurrentPesa.BadgeInStringaPesata = g_rsPese!BadgeInStringaPesata
        mCurrentPesa.ModalitaLetturaPesoAutomatica = g_rsPese!ModalitaLetturaPesoAutomatica
        mCurrentPesa.ModelloPesa = g_rsPese!Modello
        mCurrentPesa.MSGPerDoppiaPesa = g_rsPese!MSGPerDoppiaPesa
        
        mCurrentPesa.CmdSceltaPesa = nullToBlank(g_rsPese!CmdSceltaPesa)
        mCurrentPesa.cmdLetturaPeso = nullToBlank(g_rsPese!cmdLetturaPeso)
        mCurrentPesa.CmdScritturaAlibiM = nullToBlank(g_rsPese!CmdScritturaAlibiM)
        mCurrentPesa.cmdPesoStabileCode = nullToBlank(g_rsPese!cmdPesoStabileCode)
        mCurrentPesa.CmdTara = nullToBlank(g_rsPese!CmdTara)
        mCurrentPesa.CmdSceltaPesaRetOKCode = nullToBlank(g_rsPese!CmdSceltaPesaRetOKCode)
        mCurrentPesa.CmdScritturaAlibiMRetERRCode = nullToBlank(g_rsPese!CmdScritturaAlibiMRetERRCode)
        
        
        StaticCodPesa = gPesaIsolaCod
    End If
    
    
    If NoChangeRow = False Then g_rsPese.Filter = ""
    
    getCurrentPesa = mCurrentPesa
End Function


Private Function ConvToBoolean(s As String) As Boolean
    ConvToBoolean = False
    
    If (s = "") Then Exit Function
    s = UCase(s)
    If (s = "1") Then
        ConvToBoolean = True
    ElseIf (s = "SI") Then
        ConvToBoolean = True
    ElseIf (s = "VERO") Then
        ConvToBoolean = True
    ElseIf (s = "TRUE") Then
        ConvToBoolean = True
    End If
    
End Function



Public Function getCountPese() As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo getCountPeseERR
    Set rs = gDBC.Execute("select count(*) from Pese where (PCNAME='' or  PCNAME is null) and active=true")
    
    getCountPese = rs.Fields(0).value
    Exit Function
getCountPeseERR:
    
End Function


Public Sub BounsAttivoEPremioAttivo(codUtente As Long, b As Boolean, P As Boolean)
    Dim rs As ADODB.Recordset
    
    On Error GoTo BounsAttivoERR
    Set rs = gDBC.Execute("SELECT TipiAnagrafe.AttivoBonus, TipiAnagrafe.AttivoPremio " & _
                          "FROM TipiAnagrafe INNER JOIN Utenti ON TipiAnagrafe.TipoAnagrafe = Utenti.TipoAnagrafe " & _
                          "WHERE Utenti.CodUtente=" & codUtente)
    
    b = rs.Fields(0).value
    P = rs.Fields(1).value
    
    Exit Sub
BounsAttivoERR:

End Sub



Public Sub ScriviLog(TipoLog As TIPO_LOG, sApplication As String, sAzione As String, sModulo As String)
    Dim sStr As String
    
    
    sStr = Trim$(Str$(TipoLog)) & "@" & Format$(Now, "MM-DD-YYYY") & "@" & Format$(Now, "hh:mm:ss") & "@" & left$(sModulo & Space$(20), 20) & "@" & left$(sApplication & "@" & sAzione & Space$(400), 400)
    Print #mFileNumber, sStr
    'If MDIMain.StatusBar1.Panels("LOG").Picture Is Nothing And TipoLog = TPL_ERRORE Then
    '    MDIMain.Timer1.Enabled = True
    'End If
End Sub

Private Function AvviaRinominaLogFile() As String
    On Error Resume Next
    Close #mFileNumber

    If FileLen(App.Path & App.EXEName & ".Log") > 0 Then
        AvviaRinominaLogFile = gPathWrkFile & "\" & TrovaMaxFileLogFile
        Name App.Path & "\" & App.EXEName & ".Log" As AvviaRinominaLogFile
    End If
    
    Kill App.Path & "\" & App.EXEName & ".Log"
    mFileNumber = FreeFile
    Open App.Path & "\" + App.EXEName + ".LOG" For Append Shared As #mFileNumber
    If FileLen(App.Path & "\" + App.EXEName + ".LOG") = 0 Then Print #mFileNumber, Now

    
End Function


Private Function TrovaMaxFileLogFile() As String
    Dim MaxFileLogFileWRK As Long
    Static sMaxFileLogFileWRK As String
    
    If sMaxFileLogFileWRK = vbNullString Then
    
        With frmIsolaEcologicaAttiva.File1
            .Pattern = PATTERNFILELOG
            .Path = gPathWrkFile
            .Refresh
            
            MaxFileLogFileWRK = 0
            If .List(.ListCount - 1) <> "L9999999.ASC" And .List(.ListCount - 1) <> vbNullString Then
                  MaxFileLogFileWRK = Val(Mid$(.List(.ListCount - 1), 2, 7))
            End If
            
            sMaxFileLogFileWRK = "L" & Format$(Str(MaxFileLogFileWRK + 1), "0000000") & ".ASC"
        End With
    Else
            If sMaxFileLogFileWRK = "L9999999.ASC" Then sMaxFileLogFileWRK = "L0000000.ASC"
            sMaxFileLogFileWRK = "L" & Format$(Str(Val(Mid$(sMaxFileLogFileWRK, 2, 7)) + 1), "0000000") & ".ASC"
    End If
    TrovaMaxFileLogFile = sMaxFileLogFileWRK
End Function


Private Sub CancelloLogFile()
    On Error Resume Next
    
    Dim sFile As String
    Dim DateFile As Date
    Dim s As String
    
    sFile = Dir(App.Path & "\Temp\" & PATTERNFILELOG)
    Do While sFile <> vbNullString
        If FileLen(gPathWrkFile & "\" & sFile) > 0 Then
            mFileNumber = FreeFile
            Open gPathWrkFile & "\" & sFile For Input As #mFileNumber
            Line Input #mFileNumber, s
            DateFile = s
            Close #mFileNumber
            
            If DateFile <= DateAdd("d", -mGGOldFile, Now) Then
                Kill gPathWrkFile & "\" & sFile
            End If
        End If
    
        sFile = Dir
    Loop
    
    AvviaRinominaLogFile
    
End Sub

Public Property Get pesa(Index) As COMConnection
    Set pesa = mCOMconnections(Index)
End Property

Public Function getNumeroPese() As Integer
    getNumeroPese = mCOMconnections.Count
End Function

Public Sub chiudiApplicazione()
    Dim f As Form
    
    On Error Resume Next
    Set mCOMconnections = Nothing
    
    ChiudiConnessione
    
    For Each f In Forms
        Unload f
    Next
    
    On Error Resume Next
    
    
    Unload pfrmSocieta
    Unload pfrmDestinatari
    Unload pfrmTrasportatori
    Unload pfrmUtenti
    Unload pfrmUser
    Unload pfrmCodiciCER
    Unload pfrmPesata
    Unload pfrmUscita
    Unload pfrmGiacenze
    Unload pfrmTotMovimentiBonusPunti
    Unload pfrmRegistri
    Unload pfrmProtezioneDB
    Unload pfrmRegistri
    Unload pfrmFormulari
    Unload pfrmIntermediari
    Unload pfrmImpImmobiliUtenti
    Unload pfrmScuole
    
    
    
    Set pfrmScuole = Nothing
    Set pfrmSocieta = Nothing
    Set pfrmDestinatari = Nothing
    Set pfrmTrasportatori = Nothing
    Set pfrmUtenti = Nothing
    Set pfrmUser = Nothing
    Set pfrmCodiciCER = Nothing
    Set pfrmPesata = Nothing
    Set pfrmUscita = Nothing
    Set pfrmGiacenze = Nothing
    Set pfrmTotMovimentiBonusPunti = Nothing
    Set pfrmRegistri = Nothing
    Set pfrmProtezioneDB = Nothing
    Set pfrmRegistri = Nothing
    Set pfrmFormulari = Nothing
    Set pfrmIntermediari = Nothing
    Set pfrmImpImmobiliUtenti = Nothing
        
    
    Close mFileNumber
End Sub


Function ApriConnessione(Optional exclusive As Boolean = False) As Boolean
    Dim password As String
    Dim numPassword As Integer
    On Error GoTo ApriConnessioneERR
    
    If Not gDBC Is Nothing Then
        ApriConnessione = True
        Exit Function
    End If
    ApriConnessione = False
    gPathDB = GetSetting(App.Title, "DATABASE", "PATH", "")
    If gPathDB = "" Then
        gPathDB = App.Path
        SaveSetting App.Title, "DATABASE", "PATH", gPathDB
    End If
    gPathDB = gPathDB & "\IsolaEcologica.mdb"
    Set gDBC = New ADODB.Connection
    
    numPassword = 0
    With gDBC
        Do
            password = getPassword(numPassword)
            On Error Resume Next
            If exclusive Then
                .Mode = adModeShareExclusive
            End If
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gPathDB & ";Jet OLEDB:Database Password=" + password
            
            .Open
            
            If .Errors.Count > 0 Then
                If (.Errors(0).NativeError <> -124782449 And .Errors(0).Number <> 0) Then
                    On Error GoTo ApriConnessioneERR
                    Err.Raise .Errors(0).NativeError, "OPEN_DB", .Errors(0).Description
                End If
                .Errors.Clear
            End If
        Loop While password <> "" And .state <> adStateOpen
    End With
    

    gPasswordDB = password
    On Error GoTo ApriConnessioneERR
    'Provider=Microsoft.Jet.OLEDB.4.0;Password="";User ID=Admin;Data Source=C:\sviluppi\GFAmbienteEuro\StudioAmbiente2000.mdb;Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False
    If gDBC.state <> adStateOpen Then
        Err.Raise 9999, "DB OPEN", "Attenzione connessione DB mancante !"
    End If
    
    'Set gADOX.ActiveConnection = gDBC
    ApriConnessione = True
    Exit Function
ApriConnessioneERR:
    ChiudiConnessione
End Function

Public Sub ChiudiConnessione()
    'Set gADOX.ActiveConnection = Nothing
    'Set gADOX = Nothing
    If Not gDBC Is Nothing Then
        If gDBC.state = adStateOpen Then gDBC.Close
        Set gDBC = Nothing
    End If
End Sub

Public Function getMessaggioPesataLAumas400(msg As String, codBadge As String) As messaggioPesa
    'MESSAGGIO
    '&N001201L001301/2
    On Error GoTo getMessaggioPesataLAumas400ERR
    
    Dim i As Integer
    i = InStr(msg, "&")
    If i >= 1 Then
        msg = Mid$(msg, i)
        With getMessaggioPesataLAumas400
            If IsNumeric(Mid$(msg, PesaMSGLAumas400PesoNettoDa, PesaMSGLAumas400PesoNettoCaratteri)) Then
                .PesoNetto = Mid$(msg, PesaMSGLAumas400PesoNettoDa, PesaMSGLAumas400PesoNettoCaratteri) & "," & _
                         Mid$(msg, PesaMSGLAumas400PesoNettoDaDec, PesaMSGLAumas400PesoNettoDecCaratteri)
                .codBadge = codBadge
                .Data = Now
                .CodRifiuto = Space(10)
                .PesoStabile = True
            End If
        End With
    End If
    Exit Function
getMessaggioPesataLAumas400ERR:
    MsgBox "Verificare che la pesa " & gPesaIsolaCod & " sia impostata correttamente " & Err.Description
    msg = ""
    Err.Raise 9999, "Pesata Annullata"
    
End Function

Public Function getMessaggioPesataRETEL(msg As String, codBadge As String) As messaggioPesa
    'MESSAGGIO
    'hh,kk,pppppppp,uu + CR + LF
    On Error GoTo getMessaggioPesataRETELERR
    With getMessaggioPesataRETEL
        If IsNumeric(Mid$(msg, PesaMSGRETELPesoNettoDa, PesaMSGRETELPesoNettoCaratteri)) Then
            .PesoNetto = Mid$(msg, PesaMSGRETELPesoNettoDa, PesaMSGRETELPesoNettoCaratteri) & "," & _
                     Mid$(msg, PesaMSGRETELPesoNettoDaDec, PesaMSGRETELPesoNettoDecCaratteri)
            .codBadge = codBadge
            .PesoStabile = Mid$(msg, PesaMSGRETELPesoStabileDa, PesaMSGRETELPesoStabileCaratteri) = getCurrentPesa.cmdPesoStabileCode
            .Data = Now
            .CodRifiuto = Space(10)
            
        End If
    End With
    Exit Function
getMessaggioPesataRETELERR:
    MsgBox "Verificare che la pesa " & gPesaIsolaCod & " sia impostata correttamente " & Err.Description
    msg = ""
    Err.Raise 9999, "Pesata Annullata"
    
End Function



Public Function getMessaggioPesata(msg As String) As messaggioPesa
    'MESSAGGIO
    Dim peso1 As String
    Dim peso2 As String
    
    On Error GoTo getMessaggioPesataERR
    With getMessaggioPesata
        .TipoOperazione = Mid$(msg, PesaMSGTipoOpeDa, PesaMSGTipoOpeNrCaratteri)
        .Data = Mid$(msg, PesaMSGDataDa, PesaMSGDataNrCaratteri) & " " & Mid$(msg, PesaMSGOraDa, PesaMSGOraNrCaratteri)
        .Progressivo = Mid$(msg, PesaMSGProgressivoDa, PesaMSGProgressivoNrCaratteri)
        .CodRifiuto = Mid$(msg, PesaMSGCodRifiutoDa, PesaMSGCodRifiutoNrCaratteri)
        .codBadge = Mid$(msg, PesaMSGCodBadgeDa, PesaMSGCodBadgeNrCaratteri)
        .PesoStabile = True

        If getCurrentPesa.MSGPerDoppiaPesa = True Then
            peso1 = Mid$(msg, PesaMSGPesoNetto_2, PesaMSGPesoNettoNrCaratteri_2)
            peso2 = Mid$(msg, PesaMSGPesoNetto_2, PesaMSGPesoNettoNrCaratteri_2)
            .PesoNetto = IIf(getNumber(peso1) = 0, peso2, peso1)
        Else
            peso1 = Mid$(msg, PesaMSGPesoUscitaDa1, PesaMSGPesoUscitaDa1NrCaratteri1)
            peso2 = Mid$(msg, PesaMSGPesoUscitaDa2, PesaMSGPesoUscitaDa2NrCaratteri2)
            .PesoNetto = IIf(getNumber(peso1) = 0, peso2, peso1)
        End If
    End With
    
    
    Exit Function
getMessaggioPesataERR:
    MsgBox "Verificare che la pesa " & gPesaIsolaCod & " sia impostata correttamente " & Err.Description
    
    msg = ""
    Err.Raise 9999, "Pesata Annullata"
End Function

Public Function getMessaggioPesataEcoPunto(msg As String) As messaggioPesa
    'MESSAGGIO
    Dim peso1 As String
    
    With getMessaggioPesataEcoPunto
        .TipoOperazione = 0
        .Data = Mid$(msg, PesaMSGEcoDataDa, PesaMSGEcoDataNrCaratteri) & " " & Mid$(msg, PesaMSGEcoOraDa, PesaMSGEcoOraNrCaratteri)
        .Progressivo = 0
        .CodRifiuto = Mid$(msg, PesaMSGEcoCodRifiutoDa, PesaMSGEcoCodRifiutoNrCaratteri)
        .codBadge = Mid$(msg, PesaMSGEcoCodBadgeDa, PesaMSGEcoCodBadgeNrCaratteri)
        
        peso1 = Mid$(msg, PesaMSGEcoPesoNettoDa1, PesaMSGEcoPesoNettoNrCaratteri1)
        .PesoNetto = getNumber(peso1)
        .PesoStabile = True
    End With

End Function


Function nullToBlank(s As Variant)
    nullToBlank = IIf(IsNull(s), "", s)
End Function


Function toZero(s As Variant)
    
    toZero = IIf(IsNull(s), 0, s)
    If VarType(s) = vbString Then
        toZero = IIf(s = vbNullString, 0, s)
    End If
    
End Function



Sub coloraRiga(item As ListItem, color As Long)
    Dim i As Integer
    
    item.ForeColor = color
    
    For i = 1 To item.ListSubItems.Count
        item.ListSubItems(i).ForeColor = color
    Next
End Sub

Sub AllargaLeColonneLVW(lvw As ListView, AllNumCarAgg As Boolean, ParamArray NumCarAggiunti() As Variant)
    Dim i As Long
    Dim l As Long
    Dim item  As ListItem
    Dim NumCar As Integer
    
    
    For i = 0 To lvw.ColumnHeaders.Count - 1
        'Aggiungo una riga con i nomi dei titoli per allargare le colonne per
        'incaso che non vi siano record da visualizzare
        
        On Error Resume Next
        NumCar = 0
        NumCar = Abs(NumCarAggiunti(i))

        If AllNumCarAgg Then NumCar = Abs(NumCarAggiunti(0))
        If i = 0 Then
            Set item = lvw.ListItems.Add(, "TESTATACOLONNE", lvw.ColumnHeaders(i + 1).Text & String(NumCar, "W"))
        Else
            item.SubItems(i) = lvw.ColumnHeaders(i + 1).Text & String(NumCar, "W")
        End If
    
    Next

    For i = 0 To lvw.ColumnHeaders.Count - 1
        If lvw.ColumnHeaders(i + 1).width <> 0 And Sgn(NumCarAggiunti(i)) = 1 Then
            l = SendMessage(lvw.hWnd, 4126, i, -1)
        End If
    Next
    
    
    'Rimuovo la riga con le intestazioni
    lvw.ListItems.Remove ("TESTATACOLONNE")
    lvw.Visible = True
End Sub




Public Function getNumeratore(Entita As enumEntita, Optional autoCommit As Boolean = True, Optional defaultValue As Long = 1) As Long
    Dim sEntita As String
    Dim rs As ADODB.Recordset
    
    If autoCommit Then gDBC.BeginTrans
    
    getNumeratore = defaultValue
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    Set rs.ActiveConnection = gDBC
    
    If Entita = MOVIMENTI Then
        sEntita = "MOVIMENTI"
    ElseIf Entita = FATTURE Then
        sEntita = "FATTURE"
    ElseIf Entita = FORMULARIO Then
        sEntita = "FORMULARIO"
    ElseIf Entita = REGISTRO Then
        sEntita = "REGISTRO"
    ElseIf Entita = IMMOBILI Then
        sEntita = "IMMOBILI"
    ElseIf Entita = UTENTI Then
        sEntita = "UTENTI"
    ElseIf Entita = DOCUMENTO_PESATA_IN Then
        sEntita = "DOCPESA_I"
    ElseIf Entita = DOCUMENTO_PESATA_OUT Then
        sEntita = "DOCPESA_O"
    End If
    
    rs.Open ("select numeratore from contatori where entita='" & sEntita & "'")
    
    If rs.EOF Then
        gDBC.Execute ("insert into contatori (entita,numeratore) values ('" & sEntita & "'," & defaultValue & ")")
    Else
        getNumeratore = rs(0).value + 1
        gDBC.Execute ("update contatori set numeratore =" & getNumeratore & " where entita ='" & sEntita & "'")
    End If
    
    If autoCommit Then gDBC.CommitTrans
End Function


Public Function SelezionaMovimentiTotAnnuo(Anno As String, ConBonus As Boolean) As Recordset
    Dim rs As ADODB.Recordset
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    
    
    rs.Open ("SELECT CodiciCerComunali.CodCategoria, CodiciCerComunali.CodSottoCategoria, CodiciCerComunali.CodCer, Sum(Movimenti.BonusDovuto) AS SommaDiBonusDovuto,Sum(Movimenti.Quantita) AS SommaDiQuantita " & _
             "FROM CodiciCerComunali LEFT JOIN Movimenti ON (CodiciCerComunali.CodSubCer = Movimenti.CodSubCer) And (CodiciCerComunali.CodComune = Movimenti.CodComune) and (CodiciCerComunali.CodCer = Movimenti.CodCer) AND (CodiciCerComunali.CodSottoCategoria = Movimenti.CodSottoCategoria) AND (CodiciCerComunali.CodCategoria = Movimenti.CodCategoria) " & _
             "GROUP BY CodiciCerComunali.FlgBonus, CodiciCerComunali.CodCategoria, CodiciCerComunali.CodSottoCategoria, CodiciCerComunali.CodCer, Year([DataMovimento]),CodiciCerComunali.FlgAbilitatoInRicerca " & _
             "HAVING " & IIf(ConBonus, "CodiciCerComunali.FlgBonus=True", "CodiciCerComunali.FlgAbilitatoInRicerca =true") & " AND Year([DataMovimento])=" & Anno)
    
    
    Set rs.ActiveConnection = Nothing
    
    Set SelezionaMovimentiTotAnnuo = rs
    
End Function


Public Function SelezionaCodiciCer(ConBonus As Boolean) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    Set rs.ActiveConnection = gDBC
    
    
    rs.Open ("SELECT CodCategoria, CodSottoCategoria, CodCer,CodSubCer FROM CodiciCerComunali Where " & _
              IIf(ConBonus, "CodiciCerComunali.FlgBonus=True ", "FlgAbilitatoInRicerca =true"))
    
    
    Set rs.ActiveConnection = Nothing
    
    Set SelezionaCodiciCer = rs
    
End Function


Function FormatFLD(ByVal s As String, ByVal tipo As DataTypeEnum) As String
    'Formatta il campo a seconda del tipo
    Dim i As Integer

    If (s = vbNullString) Then
        FormatFLD = "NULL"
    Else
        Select Case tipo
            Case adChar, adVarWChar, adVarChar
                FormatFLD = "'" & Replace(s, "'", "''") & "'"
            Case adDate
                FormatFLD = "#" & s & "#"
            Case adBoolean
                s = UCase$(s)
                FormatFLD = IIf(s = "TRUE" Or s = "VERO", "True", "False")
            Case Else
                i = InStr(s, ",")
                If i > 0 Then Mid$(s, i) = "."
                FormatFLD = s
        End Select
    
    End If
    
    
End Function


Public Sub salvaTarga(codUtente As Long, codTarga As String, codSoggetto As String)
    On Error GoTo salvaTargaERR
    
    Dim rs As ADODB.Recordset
    Dim rsDefault As ADODB.Recordset
    Dim default As Integer
    
    Set rs = gDBC.Execute("select 1 from automezzi where codSoggetto='" & codSoggetto & "' and codutente=" & codUtente & " and targa='" & codTarga & "'")
    Set rsDefault = gDBC.Execute("select 1 from automezzi where codSoggetto='" & codSoggetto & "' and codutente=" & codUtente & " and flgDefault=true")
    default = -1
    
    If Not rsDefault.EOF Then default = 0
    
    If rs.EOF Then
        gDBC.Execute ("insert into automezzi (codutente,codSoggetto,targa,descrizione,flgDefault) " & _
                "values (" & codUtente & ",'" & codSoggetto & "','" & codTarga & "','Targa'," & default & ")")
    End If
    
    rs.Close
    rsDefault.Close
    
    Exit Sub
salvaTargaERR:
End Sub



Public Function calcolaCRC(message As String) As String
    Dim crc As Long
    Dim i As Integer
    
    For i = 1 To Len(message)
        crc = crc Xor Asc(Mid$(message, i, 1))
    Next
    
    calcolaCRC = Right$("0" & Hex(crc), 2)

End Function


Public Function salvaRiferimentiDiCaricoSulRegistro(nrRegistro As Long, colRif As Collection)
    Dim cmdq As Command
    Dim P As Parameter
    Dim i As Long
    
    On Error GoTo salvaRiferimentiDiCaricoSulRegistroERR
    salvaRiferimentiDiCaricoSulRegistro = False
    
    Set cmdq = CreateObject("ADODB.Command")
    
    cmdq.CommandText = "INSERT INTO LegameRegistroOperazioneCarico ( idRegistro, idoperazionecarico)  values(?, ? )"
    
    Set cmdq.ActiveConnection = gDBC
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    
    For i = 1 To colRif.Count
        cmdq.Parameters(0).value = nrRegistro
        cmdq.Parameters(1).value = colRif(i)
    
        cmdq.Execute
    Next

    
    
    salvaRiferimentiDiCaricoSulRegistro = True
    Exit Function
salvaRiferimentiDiCaricoSulRegistroERR:
    Err.Raise 9999, "Salva Su Registro", Err.Description
    
End Function


Public Sub AggiornaRiferimentoInDocumento(id As Long, nrDoc As Long, riferimenti As String, tipoDoc As String)
    Dim rs As ADODB.Recordset
    
    Set rs = CreateObject("ADODB.Recordset")
    
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Source = "select * from Documenti where id=" & id
    rs.ActiveConnection = gDBC
    rs.Open

    If rs.EOF Then
        rs.AddNew
        rs!nrDocument.value = nrDoc
        rs!dateDocument.value = Now
        rs!typeDocument.value = tipoDoc
        rs!Reference.value = riferimenti
        rs.Update
        id = rs!id.value
    Else
        rs!nrDocument.value = nrDoc
        rs!dateDocument.value = Now
        rs!typeDocument.value = tipoDoc
        rs!Reference.value = riferimenti
        rs.Update
    End If
    
    rs.Close
End Sub

Public Function getIdMovimentiFromDocumentId(idMov As Long) As String
    Dim rs As ADODB.Recordset
    
    Set rs = gDBC.Execute("select * from movimenti where id=" & idMov)
    
    If rs.EOF Then
        getIdMovimentiFromDocumentId = ""
    Else
        Set rs = gDBC.Execute("select * from documenti where id=" & toZero(rs!RifDocumento.value))
        If rs.EOF Then
            getIdMovimentiFromDocumentId = idMov & ","
        Else
            getIdMovimentiFromDocumentId = rs!Reference.value
        End If
    End If
End Function


Public Function creaMovimento(nrMovimento As Long, codImmobile As Long, codUtente As Long, codComune As String, codCategoria As String, codSottoCategoria As String, codCer As String, codSubCer As String, _
        DataMovimento As String, quantita As Double, Stato As String, manuale As Boolean, codUser As String, _
        BonusDovuto As Double, PuntiDovuti As Double, UnitaMisura As String, Targa As String, NumeroFormulario As String, DataFormulario As String, _
        flgForzatura As Boolean, TargaTrasportatore As String, CodTrasportatore As Long, _
        codCausale As String, flgRegistrato As Boolean, codPesa As String, RifDocumento As Long, idAlibiMemory As String) As Long
    Dim cmdq As ADODB.Command
    Dim P As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    On Error GoTo creaMovimentoERR
    
    creaMovimento = -1
    Set cmdq = CreateObject("ADODB.Command")
    
    cmdq.CommandText = "Insert into Movimenti (CodImmobile,CodUtente,CodCategoria,CodSottoCategoria, CodCer, DataMovimento, Quantita, Stato,Manuale, " & _
        "CodUser,BonusDovuto,PuntiDovuti,UnitaMisura,Targa,NumeroFormulario,DataFormulario,flgForzatura, " & _
        "TargaTrasportatore,CodTrasportatore,nrMovimento,codCausale,flgRegistrato,CodPesa,RifDocumento,IdAlibiMemory,codComune,codSubcer)" & _
        " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adWChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 30)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 6)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    
    cmdq(0).value = IIf(codImmobile = -1, Null, codImmobile)
    cmdq(1).value = IIf(codUtente = -1, Null, codUtente)
    cmdq(2).value = codCategoria
    cmdq(3).value = codSottoCategoria
    cmdq(4).value = codCer
    cmdq(5).value = DataMovimento
    cmdq(6).value = quantita
    cmdq(7).value = Stato
    cmdq(8).value = manuale
    cmdq(9).value = codUser
    cmdq(10).value = BonusDovuto
    cmdq(11).value = PuntiDovuti
    cmdq(12).value = UnitaMisura
    cmdq(13).value = IIf(Targa = "", Null, Targa)
    cmdq(14).value = IIf(NumeroFormulario = "", Null, NumeroFormulario)
    cmdq(15).value = IIf(DataFormulario = "" Or DataFormulario = "1900/01/01", Null, DataFormulario)
    cmdq(16).value = flgForzatura
    cmdq(17).value = IIf(TargaTrasportatore = "", Null, TargaTrasportatore)
    cmdq(18).value = IIf(CodTrasportatore = -1, Null, CodTrasportatore)
    
    cmdq(19).value = nrMovimento
    cmdq(20).value = codCausale
    cmdq(21).value = flgRegistrato
    cmdq(22).value = IIf(codPesa = "", Null, codPesa)
    cmdq(23).value = RifDocumento
    cmdq(24).value = idAlibiMemory
    cmdq(25).value = codComune
    cmdq(26).value = codSubCer
    
    Set cmdq.ActiveConnection = gDBC
    
    cmdq.Execute

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    Set rs.ActiveConnection = gDBC
        
    rs.Source = "select id from Movimenti where nrMovimento = " & nrMovimento & " and DataMovimento=#" & Format$(DataMovimento, "yyyy\-mm\-dd hh\:mm\:ss") & "#"
    rs.Open
        
    creaMovimento = rs(0).value
        
    Set rs.ActiveConnection = Nothing
    rs.Close
    Set rs = Nothing
    
    Set cmdq.ActiveConnection = Nothing
    Set cmdq = Nothing
    

    Exit Function
creaMovimentoERR:
    Err.Raise 999, "CreaMovimento", Err.Description
End Function


Public Function creaVoceDiRegistro(pNrRegistro As Long, pflgOperazione As String, pdataOperazione As String, pnrFormulario As String, pdataFormulario As String, _
             pcodCategoria As String, pcodSottoCategoria As String, pcodCer As String, pcodTipoOpeSmaltRecupero As String, _
             punitaMisura As String, pquantita As Double, pImpiantoiProduzione As String, pidIntermediario As Long, pNote As String, dataIns As Date, codUser As String, classiDiPericolosita As String) As Long
    Dim cmdq As Command
    Dim P As Parameter
    Dim rs As Recordset
    
    
    On Error GoTo salvaSuRegistroERR
    creaVoceDiRegistro = -1
    
    Set cmdq = CreateObject("ADODB.Command")
    
    cmdq.CommandText = "INSERT INTO Registro ( nrRegistro, flgOperazione, dataOperazione, nrFormulario, dataFormulario, " & _
             " codCategoria, codSottocategoria, codCer, codTipoOpeSmaltRecupero, unitaMisura, quantita, ImpiantoiProduzione, idIntermediario, [note], flgStampato,DataIns,CodUser,classiPericolosita )" & _
             " values(?, ?, ?, ?, ?, ?, ?, ?,?,?, ?, ?, ?, ?, ?,?,?,?)"
             

    Set cmdq.ActiveConnection = gDBC

    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 15)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 3)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 2
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    cmdq.Parameters.Append cmdq.CreateParameter(, adBoolean, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 255)
    
    cmdq.Parameters(0).value = pNrRegistro
    cmdq.Parameters(1).value = pflgOperazione
    cmdq.Parameters(2).value = Format$(pdataOperazione, "yyyy-MM-dd")
    If pnrFormulario = "" Then
        cmdq.Parameters(3).value = Null
        cmdq.Parameters(4).value = Null
    Else
        cmdq.Parameters(3).value = pnrFormulario
        cmdq.Parameters(4).value = pdataFormulario
    End If
    cmdq.Parameters(5).value = pcodCategoria
    cmdq.Parameters(6).value = pcodSottoCategoria
    cmdq.Parameters(7).value = pcodCer
    
    cmdq.Parameters(8).value = IIf(pcodTipoOpeSmaltRecupero = "", Null, pcodTipoOpeSmaltRecupero)
    cmdq.Parameters(9).value = punitaMisura
    cmdq.Parameters(10).value = pquantita
    
    cmdq.Parameters(11).value = IIf(pImpiantoiProduzione = "", Null, pImpiantoiProduzione)
    cmdq.Parameters(12).value = IIf(pidIntermediario = 0, Null, pidIntermediario)
    cmdq.Parameters(13).value = pNote
    cmdq.Parameters(14).value = False
    cmdq.Parameters(15).value = dataIns
    cmdq.Parameters(16).value = codUser
    cmdq.Parameters(17).value = classiDiPericolosita
    
    cmdq.Execute
    
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    
    Set rs.ActiveConnection = gDBC
        
    rs.Source = "select idregistro from Registro where nrRegistro = " & pNrRegistro & " and dataOperazione=#" & Format$(pdataOperazione, "yyyy-MM-dd") & "#"
    rs.Open
    creaVoceDiRegistro = rs(0).value
        
    Set rs.ActiveConnection = Nothing
    rs.Close
    Set rs = Nothing
    
        
    Set cmdq.ActiveConnection = Nothing
    Set cmdq = Nothing

    
    Exit Function
salvaSuRegistroERR:
    Err.Raise 9999, "Salva Su Registro", Err.Description
    
 
End Function

Public Function creaLotto(pidRegistro As Long, _
        pcodCategoria As String, _
        pcodSottoCategoria As String, _
        pcodCer As String, _
        pQtaIngresso As Double, _
        pQtaUscita As Double, _
        pTipoLotto As String, _
        pStatoLotto As String, _
        pDataMod As Date, _
        pCodUserMod As String, _
        pDataIns As Date, _
        pCodUserIns As String) As String
        
    Dim cmdq As ADODB.Command
    Dim P As ADODB.Parameter
    Dim codLotto As String
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandText = "INSERT INTO Lotti ( CodLotto,  CodCategoria, CodSottoCategoria, CodCer, " & _
                       "QtaIngresso, QtaUscita,Tipo,Stato, DataMod, CodUserMod, DataIns, CodUserIns ) " & _
                       "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"


    Set cmdq.ActiveConnection = gDBC
 
    codLotto = pidRegistro & "D" & Format$(Now, "DDMMYY") & "H" & Format$(Now, "HHNNSS")
    
    creaLotto = codLotto
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 2)
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    Set P = cmdq.CreateParameter(, adDouble, adParamInput)
    P.NumericScale = 4
    P.Precision = 6

    cmdq.Parameters.Append P
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 1)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 1)
    
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    cmdq(0).value = codLotto
    cmdq(1).value = pcodCategoria
    cmdq(2).value = pcodSottoCategoria
    cmdq(3).value = pcodCer
    cmdq(4).value = pQtaIngresso
    cmdq(5).value = pQtaUscita
    cmdq(6).value = pTipoLotto
    cmdq(7).value = pStatoLotto
    cmdq(8).value = pDataMod
    cmdq(9).value = pCodUserMod
    cmdq(10).value = pDataIns
    cmdq(11).value = pCodUserIns

    cmdq.Execute
    
    Set cmdq = CreateObject("ADODB.Command")
    cmdq.CommandText = "INSERT INTO LegameLottoRegistrazioneCarico ( idRegistro,  codLotto ,descrizione, DataIns, CodUserIns ) " & _
                       "VALUES (?,?,?,?,?)"
    
    Set cmdq.ActiveConnection = gDBC
    cmdq.Parameters.Append cmdq.CreateParameter(, adInteger, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 50)
    cmdq.Parameters.Append cmdq.CreateParameter(, adDate, adParamInput)
    cmdq.Parameters.Append cmdq.CreateParameter(, adVarChar, adParamInput, 10)
    
    cmdq(0).value = pidRegistro
    cmdq(1).value = codLotto
    cmdq(2).value = "legame"
    cmdq(3).value = Format$(pDataIns, "yyyy-MM-dd")
    cmdq(4).value = pCodUserIns

    cmdq.Execute
        
End Function

Public Function findCap(codistat As String) As String
    If DevEnvMDB.rscmdComune.state = adStateOpen Then
        DevEnvMDB.rscmdComune.Close
    End If
    DevEnvMDB.cmdComune codistat
    
    findCap = ""
    
    If Not DevEnvMDB.rscmdComune.EOF Then findCap = DevEnvMDB.rscmdComune!Cap
End Function



Public Sub generazioneRegistrazioneDiCarico(Optional force As Boolean = False)
    Dim rs As Recordset
    Dim i As Long
    Dim rs1 As Recordset
    Dim rsDettaglio As Recordset
    Dim sWhere As String
    Dim idRegistro As Long
    Dim nrRegistro As Long
    Dim DataMovimento As Date
    Dim colMov As New Collection
    
    On Error GoTo generazioneRegistrazioneDiCaricoERR
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.LockType = adLockBatchOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    
    'selezione dell'ultima operazione di accorpamento
    rs.Source = "select DataCarico,DataProssimoCarico,DataUltimoMovimento,idUltimoRegistroCarico from DataGenerazioneCarico"
    Set rs.ActiveConnection = gDBC
    gDBC.BeginTrans
    
    
    rs.Open
    If Not rs.EOF Then
        If force = False And Format$(Now, "yyyy-MM-dd") < Format$(rs!DataProssimoCarico, "yyyy-MM-dd") Then
            rs.Close
            Set rs.ActiveConnection = Nothing
            
            gDBC.RollbackTrans
            Exit Sub
        End If
        sWhere = " and DataMovimento > #" & Format$(rs!DataUltimoMovimento, "yyyy-MM-dd") & " 00:00:00# " & _
                 " and DataMovimento <= #" & Format$(rs!DataProssimoCarico, "yyyy-MM-dd") & " 23:59:59# "
    End If
    
    
    'selezione dei movimenti d'accorpare
    Set rs1 = CreateObject("ADODB.Recordset")
    rs1.LockType = adLockBatchOptimistic
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenStatic
    
    Set rsDettaglio = CreateObject("ADODB.Recordset")
    rsDettaglio.LockType = adLockBatchOptimistic
    rsDettaglio.CursorLocation = adUseClient
    rsDettaglio.CursorType = adOpenStatic
    
    
    rs1.Source = "SELECT Format$(DataMovimento,'yyyy-mm-dd') AS  DataMov, Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer, Movimenti.UnitaMisura, Max(Movimenti.DataMovimento) AS DataMovimento, Sum(Movimenti.Quantita) AS Quantita " & _
                " From Movimenti " & _
                " WHERE ( Movimenti.NumeroFormulario<>'' Or Movimenti.NumeroFormulario Is Null)  AND Movimenti.flgRegistrato=false and Movimenti.Stato<>'2' AND (Movimenti.CodCausale='00' " & _
                " or  Movimenti.CodCausale='" & gCodCausaleMovManuale & "' or  Movimenti.CodCausale='" & gCodCausaleMovEcoPuntoIN & "')" & sWhere & _
                " GROUP BY  Format$(DataMovimento,'yyyy-mm-dd') , Movimenti.CodCategoria, Movimenti.CodSottoCategoria, Movimenti.CodCer, Movimenti.UnitaMisura " & _
                " ORDER BY 1 asc"

 
    rsDettaglio.Source = "SELECT Format$(DataMovimento,'yyyy-mm-dd') AS  DataMov,Movimenti.id,codCategoria,codSottoCategoria, codCer" & _
            " from Movimenti Where ( Movimenti.NumeroFormulario<>'' Or Movimenti.NumeroFormulario Is Null) AND Movimenti.flgRegistrato=false and Movimenti.Stato<>'2' AND " & _
            " (Movimenti.CodCausale='00' or  Movimenti.CodCausale='" & gCodCausaleMovManuale & "' or  Movimenti.CodCausale='" & gCodCausaleMovEcoPuntoIN & "')" & _
             sWhere

    Set rs1.ActiveConnection = gDBC
    Set rsDettaglio.ActiveConnection = gDBC
    
    rs1.Open
    rsDettaglio.Open
    
    Do While Not rs1.EOF
        nrRegistro = getNumeratore(REGISTRO, False)
        idRegistro = creaVoceDiRegistro(nrRegistro, "C", Format$(rs1!DataMovimento, "yyyy-MM-dd"), "", "", rs1!codCategoria, rs1!codSottoCategoria, rs1!codCer, "" _
                            , rs1!UnitaMisura, rs1!quantita, "", 0, "", Now, "Admin", "")
                            
        creaLotto idRegistro, rs1!codCategoria, rs1!codSottoCategoria, rs1!codCer, rs1!quantita, 0, "I", 0, _
                     Now, "Admin", rs1!DataMov, "Admin"
        
        'prendo il movimento più alto come data
        If DataMovimento < rs1!DataMovimento Then DataMovimento = rs1!DataMovimento
        
        rsDettaglio.Filter = ""
        rsDettaglio.Filter = "DataMov='" & rs1!DataMov & "' and codCategoria = '" & rs1!codCategoria & _
            "' and codSottoCategoria='" & rs1!codSottoCategoria & _
            "' and codCer='" & rs1!codCer & "'"
        
        Do While Not rsDettaglio.EOF
            colMov.Add rsDettaglio!id.value
            rsDettaglio.MoveNext
        Loop
        
        salvaRiferimentiDiCaricoSulRegistro idRegistro, colMov
        
        aggiornaMovimentiComeRegistrati colMov
        
        For i = colMov.Count To 1 Step -1
            colMov.Remove i
        Next
        
        
        
        rs1.MoveNext
    Loop
    
    If rs.EOF Then
        rs.AddNew
    End If
    
    rs!DataCarico.value = Format$(Now, "yyyy-MM-dd")
    If Not force Then rs!DataProssimoCarico.value = Format$(DateAdd("d", gRegCaricoSuRegistroOgni, rs!DataCarico.value), "yyyy-MM-dd")
    If DataMovimento = "0.00.00" Then DataMovimento = "1999-01-01"
    rs!DataUltimoMovimento.value = Format$(DataMovimento, "yyyy-MM-dd")
    rs!idUltimoRegistroCarico.value = idRegistro
    rs.UpdateBatch adAffectAllChapters
    
    rs.Close
    Set rs.ActiveConnection = Nothing
    
    rsDettaglio.Close
    Set rsDettaglio.ActiveConnection = Nothing
    
    rs1.Close
    Set rs1.ActiveConnection = Nothing
    
    gDBC.CommitTrans
    Exit Sub
generazioneRegistrazioneDiCaricoERR:
    
    gDBC.RollbackTrans
    
    Err.Raise 9999, "RegistroCarico", Err.Description
End Sub


Private Sub aggiornaMovimentiComeRegistrati(colMov As Collection)
    Dim i As Integer
    Dim qryUPD As Command
    
    Set qryUPD = CreateObject("ADODB.Command")
    qryUPD.CommandText = "update  Movimenti Set flgRegistrato = true  where id = ?"
    Set qryUPD.ActiveConnection = gDBC
    qryUPD.Parameters.Refresh
    
    For i = 1 To colMov.Count
        qryUPD.Parameters(0).value = colMov(i)
        qryUPD.Execute
    Next
    
    Set qryUPD.ActiveConnection = Nothing
    Set qryUPD = Nothing
End Sub

Private Function countNumChar(s As String, sChar As String) As Long
    Dim i As Integer
    
    countNumChar = 0
    i = 1
    Do
        i = InStr(i, s, sChar)
        If i > 0 Then
            countNumChar = countNumChar + 1
            i = i + 1
        End If
    Loop While i > 0
End Function


Public Function getNumber(ByVal s As String) As Double
    Dim sN As String
    Dim n As Double
    Dim decimalSeparator As String
    Dim thousandSeparator As String
    
    Dim idecimalSeparator As Long
    Dim ithousandSeparator As Long
    
    Dim iCountDecimalSeparator As Long
    Dim iCountThousandSeparator As Long
    
    
    n = 1.2
    sN = n
    
    decimalSeparator = Mid$(sN, 2, 1)
    thousandSeparator = IIf(decimalSeparator = ".", ",", ".")
    
    iCountDecimalSeparator = countNumChar(s, decimalSeparator)
    iCountThousandSeparator = countNumChar(s, thousandSeparator)
    
    
    If iCountThousandSeparator > 1 Then
        s = Replace(s, thousandSeparator, vbNullString)
    End If
    
    If iCountDecimalSeparator > 1 Then
        s = Replace(s, decimalSeparator, vbNullString)
    End If
    
    idecimalSeparator = InStrRev(s, decimalSeparator)
    ithousandSeparator = InStrRev(s, thousandSeparator)
    
    If idecimalSeparator > ithousandSeparator Then
        n = s
    ElseIf idecimalSeparator < ithousandSeparator Then
        n = Replace(Replace(s, decimalSeparator, ""), thousandSeparator, decimalSeparator)
    Else
        n = s
    End If
    
    getNumber = n
End Function


Private Sub caricaTracciati()
    Dim f As New FileSystemObject
    Dim txtFile As TextStream
    Dim s As String
    Dim sCol() As String
    Dim fieldSize() As String
    Dim typef As FIELD_TYPE_ENUM
    Dim fieldPrecision As Long
    
    Set txtFile = f.OpenTextFile(App.Path & "\TracciatoImmUte.ini", ForReading)
    Set colTracciati = New Tracciati
    
    s = txtFile.ReadLine
    Do While Not txtFile.AtEndOfStream
        s = txtFile.ReadLine
        sCol = Split(s, ";")
        fieldSize = Split(sCol(3), ",")
        fieldPrecision = 0
        If UBound(fieldSize) > 0 Then
            fieldPrecision = fieldSize(1)
        End If
        
        If UCase$(sCol(2)) = "CHAR" Then
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_CHAR
        ElseIf UCase$(sCol(2)) = "BOOLEAN" Then
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_BOOL
        Else
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_NUM
        End If
        colTracciati.Add "IMMUTE", sCol(0), sCol(1), typef, CLng(fieldSize(0)), fieldPrecision, sCol(4), sCol(5), sCol(6), sCol(7), sCol(8)
    Loop
    
    txtFile.Close
    Set txtFile = Nothing
    Set f = Nothing
    
    
    Set txtFile = f.OpenTextFile(App.Path & "\TracciatoUteBadge.ini", ForReading)
    
    s = txtFile.ReadLine
    Do While Not txtFile.AtEndOfStream
        s = txtFile.ReadLine
        sCol = Split(s, ";")
        fieldSize = Split(sCol(3), ",")
        fieldPrecision = 0
        If UBound(fieldSize) > 0 Then
            fieldPrecision = fieldSize(1)
        End If
        
        If UCase$(sCol(2)) = "CHAR" Then
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_CHAR
        ElseIf UCase$(sCol(2)) = "BOOLEAN" Then
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_BOOL
        Else
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_NUM
        End If
        colTracciati.Add "UTEBADGE", sCol(0), sCol(1), typef, CLng(fieldSize(0)), fieldPrecision, sCol(4), sCol(5), sCol(6), sCol(7), sCol(8)
    Loop
    
    txtFile.Close
    Set txtFile = Nothing
    Set f = Nothing
    
    Set txtFile = f.OpenTextFile(App.Path & "\TracciatoLegameBarcodeBadge.ini", ForReading)
    
    s = txtFile.ReadLine
    Do While Not txtFile.AtEndOfStream
        s = txtFile.ReadLine
        sCol = Split(s, ";")
        fieldSize = Split(sCol(3), ",")
        fieldPrecision = 0
        If UBound(fieldSize) > 0 Then
            fieldPrecision = fieldSize(1)
        End If
        
        If UCase$(sCol(2)) = "CHAR" Then
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_CHAR
        ElseIf UCase$(sCol(2)) = "BOOLEAN" Then
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_BOOL
        Else
            typef = FIELD_TYPE_ENUM.FIELD_TYPE_NUM
        End If
        colTracciati.Add "LEGBADGE", sCol(0), sCol(1), typef, CLng(fieldSize(0)), fieldPrecision, sCol(4), sCol(5), sCol(6), sCol(7), sCol(8)
    Loop
    
    txtFile.Close
    Set txtFile = Nothing
    Set f = Nothing
    
    
End Sub



Public Function getCodiceUtenteDaBadge(codBadge As String) As Long
    Dim rs As Recordset
    
    getCodiceUtenteDaBadge = 0
    Set rs = gDBC.Execute("select codUtente from badge where codBadge='" & Trim$(codBadge) & "'")
    If Not rs.EOF Then getCodiceUtenteDaBadge = rs(0).value
    
    rs.Close
    Set rs = Nothing
End Function

Public Sub ComprimiDB()
    Dim i As Long
    
    On Error GoTo ComprimiDBERR

    Screen.MousePointer = vbHourglass

    For i = 1 To DevEnvMDB.Connections.Count
        DevEnvMDB.Connections(i).Close
    Next
    
    DoEvents
    
    ChiudiConnessione
    If Dir(App.Path & "\Comp.mdb") <> "" Then Kill App.Path & "\Comp.mdb"
    
    gJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gPathDB & ";Jet OLEDB:Database Password=" + gPasswordDB _
    , "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Comp.mdb"
    
    Kill gPathDB
    Name App.Path & "\Comp.mdb" As gPathDB
    
    
    'reimposta password
    ApriConnessione True
    
    On Error Resume Next
    gDBC.Execute "ALTER Database Password " & gPasswordDB & " " & gPasswordDB
    If gDBC.Errors.Count > 0 Then
        gDBC.Errors.Clear
        On Error GoTo ComprimiDBERR
        gDBC.Execute "ALTER Database Password " & gPasswordDB & " null"
    End If

    
    ChiudiConnessione
    
    ApriConnessione
  
    
    For i = 1 To DevEnvMDB.Connections.Count
        DevEnvMDB.Connections(i).Open
    Next
    
    
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "Compressione DataBase effettuata con successo!!", vbExclamation + vbOKOnly, App.Title
    
    Exit Sub
ComprimiDBERR:
    ApriConnessione
    
    For i = 1 To DevEnvMDB.Connections.Count
        DevEnvMDB.Connections(i).Open
    Next
    
    
    MsgBox "CompimiDB:" & Err.Description, vbCritical + vbOKOnly, App.Title
    Screen.MousePointer = vbDefault
End Sub

Public Sub creaImmobileUtenteEBadge(codBadge As String)
    Dim rs As ADODB.Recordset
    Dim codUtente As Long
    Dim codImmobile As Long
    Dim codComune As String
    
    On Error GoTo creaImmobileUtenteEBadgeERR
    Set rs = gDBC.Execute("select 1 from badge where codBadge ='" & codBadge & "'")
    If rs.EOF Then
    
        Dim rsComune As ADODB.Recordset
        Set rsComune = GetCurrentComune
        If rsComune Is Nothing Then Exit Sub
        
        codComune = rsComune!codistat
    
        'creo utente ed immobile e badge
        gDBC.BeginTrans
        codUtente = getNumeratore(UTENTI, False)
        
'        MsgBox "CODBADGE :" & codBadge
'        ScriviLog TPL_INFORAMITIVO, "insert into immobili (codImmobile,codUtente,Tipo,codComune,Indirizzo,CodTariffa,TipoUtenza,flgUtenzaDomesticaResidente,DataInizioValidita,DataFineValidita) " & _
'                      "values (" & codImmobile & "," & codUtente & ",'P','" & codComune & "','.','Ud 00','D',1,date(), #9999-01-01#)", "sss", "aaaa"
        
'        ScriviLog TPL_INFORAMITIVO, "insert into utenti (codutente,codanagrafe,tipoanagrafe,nome,cognome,ragionesociale,IndirizzoResidenza,NumCivicoResidenza,codComuneResidenza,tipodomicilio) " & _
'                      "values (" & codUtente & "," & codUtente & ",'F','utente" & codUtente & "','.','.','.',0,'" & codComune & "','R')", "sss", "aaaa"
                      
'        ScriviLog TPL_INFORAMITIVO, "insert into badge (codBadge,codUtente,tipoanagrafe,progressivobadge,dataBadgeInizio,dataBadgeFine,codcomune,codImmobile,tipoutenza,[default]) " & _
'                      "values (" & codBadge & "," & codUtente & ",'F',0,date(),#9999-01-01#,'" & codComune & "',0,'F',1)", "sss", "aaaa"
        
        If codUtente > 0 Then
            gDBC.Execute ("insert into utenti (codutente,codanagrafe,tipoanagrafe,nome,cognome,ragionesociale,IndirizzoResidenza,NumCivicoResidenza,codComuneResidenza,tipodomicilio) " & _
                          "values (" & codUtente & "," & codUtente & ",'F','utente" & codUtente & "','.','.','',0,'" & codComune & "','R')")
                          
            codImmobile = getNumeratore(IMMOBILI, False)
            If codImmobile > 0 Then
                gDBC.Execute ("insert into immobili (codImmobile,codUtente,Tipo,codComune,Indirizzo,numCivico,CodTariffa,TipoUtenza,flgUtenzaDomesticaResidente,DataInizioValidita,DataFineValidita) " & _
                              "values (" & codImmobile & "," & codUtente & ",'P','" & codComune & "','.',0,'Ud 00','D',1,date(), #9999-01-01#)")
                              
                              
                gDBC.Execute ("insert into badge (codBadge,codUtente,tipoanagrafe,progressivobadge,dataBadgeInizio,dataBadgeFine,codcomune,codImmobile,tipoutenza,[default]) " & _
                              "values ('" & codBadge & "'," & codUtente & ",'F',0,date(),#9999-01-01#,'" & codComune & "'," & codImmobile & ",'F',1)")
            Else
                Err.Raise 99999, "creaImmobileUtenteEBadge", "Attenzione non trovato numeratore per immobile"
            End If
        Else
            Err.Raise 99999, "creaImmobileUtenteEBadge", "Attenzione non trovato numeratore per utente"
        End If
                      
        gDBC.CommitTrans
        
        
    End If
    
    Exit Sub
creaImmobileUtenteEBadgeERR:
    gDBC.RollbackTrans
    MsgBox "CreaImmobileUtenteBadge:" & Err.Description
End Sub


Public Function GetTipiAnagrafe() As String
    GetTipiAnagrafe = Choose(gPuntiAssegnatiPer, "I - Istituto", "S - Scuola", "C - Classe", "Z - Sezione", "E - Studente")
End Function


'inizializza le pese
Public Sub InitPese(splash As Boolean, pSplash As Form)
    Dim i As Long
    Dim oldId As Long
    
    On Error GoTo InitPeseERR
    'cancello tutte le com create
    For i = mCOMconnections.Count To 1 Step -1
        mCOMconnections.Remove i
    Next
     
     
    For i = MDIMain.MSComm1.UBound To MDIMain.MSComm1.LBound + 1 Step -1
        Unload MDIMain.MSComm1(i)
    Next
     
    'aggiungo le com collegate alle pese
    g_rsPese.MoveFirst
    g_rsPese.Sort = "ComID"
    i = 0

    Do While Not g_rsPese.EOF
        If nullToBlank(g_rsPese!PCNAME) = vbNullString And g_rsPese!Active Then
            Set mCOM = mCOMconnections.Add(g_rsPese!codPesa)
            mCOM.BaudeRate = g_rsPese!BaudeRate
            mCOM.ComID = g_rsPese!ComID
            mCOM.DataBits = g_rsPese!DataBits
            mCOM.ParityBit = g_rsPese!ParityBit
            mCOM.StopBit = g_rsPese!StopBit
            mCOM.InBufferSize = 1024
            mCOM.CDHolding = g_rsPese!CheckPort
            mCOM.DSRHolding = g_rsPese!CheckPort
            mCOM.HandShaking = IIf(g_rsPese!XonXoff, 1, 0)
    
            If splash Then pSplash.setInfo "Inizializza Pesa " & i & " . . ."
    
            
            mCOM.StartChar = IIf(left$(g_rsPese!StartChar, 1) = "N", vbNullString, g_rsPese!StartChar)
            mCOM.message.STARCHAR_ON = IIf(mCOM.StartChar = vbNullString, False, True)
           
            mCOM.message.CheckSumON = g_rsPese!CheckSum
            mCOM.message.LineFeedON = False
            
 
            mCOM.STX = nullToBlank(g_rsPese!STX)
            mCOM.ETX = nullToBlank(g_rsPese!ETX)
            
            mCOM.message.STXETX_ON = mCOM.STX <> vbNullString
            mCOM.message.CarriageReturnON = IIf(g_rsPese!CR = True, True, False)
    
            If oldId <> g_rsPese!ComID Then
                If (mCOM.Init(MDIMain.MSComm1) = False) Then ScriviLog TPL_ERRORE, App.Title, "INITCOM", "Comunicazione con la Pesa non avvenuta"
                oldId = g_rsPese!ComID
            Else
                Set mCOM.COMControl = getCOMByID_COM(oldId)
            End If

            MDIMain.PesaCollegata i + 1, g_rsPese!Descrizione
            
            i = i + 1
        End If
        g_rsPese.MoveNext
    Loop
    
    'se esiste il comando di selezione pesa lo invio
    InvioComandoAllaBilancia cmd_CambioCanale, False
    
    
    letturaStrutturaMessaggioPesa
InitPeseERR:
   
End Sub

Public Function InvioComandoAllaBilancia(ComandiBilancia As ComandiBilanciaEnum, Notify As Boolean, Optional AlibiMemory As String, Optional NoChangeCurrentBalance As Boolean = False) As Boolean
    Dim s As String
    Dim sRet As String
    Dim t As Long
    Dim p_currentPesa As currentPesa
    
    InvioComandoAllaBilancia = False
    
    p_currentPesa = getCurrentPesa(NoChangeCurrentBalance)
    
    pesa(gPesaIsolaCod).ClearBuffer
    
    Select Case ComandiBilancia
        Case cmd_RichiestaPesa
           If p_currentPesa.cmdLetturaPeso = vbNullString Then Exit Function
           
           
           If pesa(gPesaIsolaCod).message.STXETX_ON Then
                pesa(gPesaIsolaCod).SendMessage pesa(gPesaIsolaCod).STX & p_currentPesa.cmdLetturaPeso & pesa(gPesaIsolaCod).ETX
           Else
                pesa(gPesaIsolaCod).SendMessage p_currentPesa.cmdLetturaPeso + vbCrLf
           End If
           Sleep 200
        Case cmd_ScritturaAlibi
        
            If p_currentPesa.CmdScritturaAlibiM = vbNullString Then Exit Function
         
            If pesa(gPesaIsolaCod).message.STXETX_ON Then
                 pesa(gPesaIsolaCod).SendMessage pesa(gPesaIsolaCod).STX & p_currentPesa.CmdScritturaAlibiM & pesa(gPesaIsolaCod).ETX
            Else
                 pesa(gPesaIsolaCod).SendMessage p_currentPesa.CmdScritturaAlibiM + vbCrLf
            End If
            Sleep 200
            
            t = Timer
            s = ""
            AlibiMemory = "NOALIBI"
            Do While Timer < t + 10
                
                Select Case pesa(gPesaIsolaCod).ReadMessage(s)
                    Case ErroreRicevuto
                        Debug.Print s
                    Case InLettura
                    Case LetturaTerminata
                        Debug.Print Now & " - " & s
                        Select Case p_currentPesa.ModelloPesa
                            Case "RETEL"
                                If left$(s, 3) = p_currentPesa.CmdScritturaAlibiM Then
                                    AlibiMemory = Trim$(Mid$(s, 37, 12))
                                    If AlibiMemory = p_currentPesa.CmdScritturaAlibiMRetERRCode Then AlibiMemory = "ERRALIBI"
                                    
                                    Exit Do
                                End If
                        End Select
                        
                End Select
                DoEvents
            Loop
                 
            If Notify And AlibiMemory = vbNullString Or AlibiMemory = "ERRALIBI" Or AlibiMemory = "NOALIBI" Then
                'MsgBox "Attenzione l'ALIBI MEMORY non è stata scritta correttamente, provare a salvare nuovamente", vbCritical + vbOKOnly, App.Title
                InvioComandoAllaBilancia = False
                Exit Function
            End If
            
            
            DoEvents
        Case cmd_CambioCanale
            
            
            If p_currentPesa.CmdSceltaPesa = vbNullString Then Exit Function
            
            If pesa(gPesaIsolaCod).message.STXETX_ON Then
                 pesa(gPesaIsolaCod).SendMessage pesa(gPesaIsolaCod).STX & p_currentPesa.CmdSceltaPesa & pesa(gPesaIsolaCod).ETX
            Else
                 pesa(gPesaIsolaCod).SendMessage p_currentPesa.CmdSceltaPesa & vbCrLf
            End If
            Sleep 200
            
            s = ""
            t = Timer
            Do While Timer < t + 5
                 
                Select Case pesa(gPesaIsolaCod).ReadMessage(s)
                    Case ErroreRicevuto
                    Case InLettura
                    Case LetturaTerminata
                        If s = p_currentPesa.CmdSceltaPesaRetOKCode Then Exit Do
                End Select
            Loop
             
            If s = p_currentPesa.CmdSceltaPesaRetOKCode Then
                If Notify Then MsgBox "La pesa è stata selezionata"
            Else
                If Notify Then MsgBox "La pesa è non è stata selezionata verificare le connessioni"
                InvioComandoAllaBilancia = False
                Exit Function
            End If

        
        Case cmd_Tara
    
    End Select
    
    InvioComandoAllaBilancia = True
End Function



Function getCOMByID_COM(ComID As Long) As MSComm
    Dim i As Integer
    
    For i = 1 To mCOMconnections.Count
        If Not mCOMconnections(i).COMControl Is Nothing Then
            If mCOMconnections(i).COMControl.CommPort = ComID Then
                Set getCOMByID_COM = mCOMconnections(i).COMControl
                Exit For
            End If
        End If
    Next
End Function

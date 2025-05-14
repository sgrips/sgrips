Attribute VB_Name = "ModMain"
Option Explicit

Public gGiorniScandenzaDemo As Integer
Public gUnloadMdi As Boolean
Public pGrafici(9) As frmGrafici

Public Const SEPARATORE = "&"
Public Const RIDUZIONE_SINGLE = "@Single@"
Public Const FORMAT6DEC = "###,###,###,###,##0.000000"
Public Const FORMAT5DEC = "###,###,###,###,##0.00000"
Public Const FORMAT4DEC = "###,###,###,###,##0.0000"
Public Const FORMAT3DEC = "###,###,###,###,##0.000"
Public Const FORMATVALUTA = "###,###,###,###,##0.00"
Public Const FORMATSTANDARD = "###,###,###,###,##0"
Public Const EM_LINESCROLL = &HB6

Public Const VALUTA_EURO = "€"
Public Const VALUTA_LIRE = "£"

Public Const ALL_COMUNI = "ALLCOM"
Public Const ID_ERR_PERSONALE = 10000
Public Const SEZ_ULTIMI_VALORI = "UltimiValori"
Public Const SEZ_FILTRI = "Filtri"
Public Const KEY_ULTIMI_VALORI_COMUNESEL = "Comune selezionato"
Public Const KEY_ULTIMI_VALORI_VALUTA = "Valuta"
Public Const KEY_FILTRI_DATADA = "DataDa"
Public Const KEY_FILTRI_DATAA = "DataA"
Public Const KEY_FILTRI_COMUNE = "Comune"

'Icone
Public Const ICON_FACTORY = "Icon_Factory"
Public Const ICON_PF = "Icon_PF"
Public Const ICON_PV = "Icon_PV"
Public Const ICON_MAN = "Icon_Man"
Public Const ICON_ANAGRAFICA = "Icon_Ana"
Public Const ICON_DATIRACCOLTA = "Icon_Ana"
Public Const ICON_DATITECNICI = "Icon_DatiTecnici"
Public Const ICON_NULL = "Icon_Null"
Public Const ICON_COSTI = "Icon_Costi"
Public Const ICON_AMMORTAMENTO = "Icon_Ammortamento"
Public Const ICON_PERSONALE = "Icon_Personale"
Public Const ICON_DOMESTICA = "Icon_Domestica"
Public Const ICON_ATTIVITA = "Icon_Attivita"
Public Const ICON_DOMESTICA_ATTIVITA = "Icon_Domestica_Attivita"


Public gPasswordDB As String
Public gPathDB As String
Public gADOX As New adox.Catalog
Public gDBC As ADODB.Connection
Public gJRO As JRO.JetEngine
Public gErrore As Boolean
Public gComuneAttivo As String
Public gWhereSuComune As String
Public gWhereFiltroSuComune As String
Public gcmdAnnulla As Boolean
Public gValutaDefault As String
Public gInfoTxt() As String
Public gidxInfo As Integer

'Propieta scritte in un TAG
Public Const TAG_TIPO_FLD = SEPARATORE & "Tipo"
Public Const TAG_SIZE_FLD = SEPARATORE & "Size"
Public Const TAG_UD_FLD = SEPARATORE & "UD"
Public Const TAG_UA_FLD = SEPARATORE & "UA"

 
'Enum Tipi Dati da caricare nel ListVIEW
Public Enum TipoDatoLvwEnum
    Lvw_AnagraficaComuni = 1
    Lvw_CostiComunali = 2
    Lvw_TipiAmmortamenti = 3
    Lvw_Personale = 4
    Lvw_DatiTecnici = 5
    Lvw_AnagraficaRaccolta = 6
End Enum

'Campi Calcolati
Public Const FORECOLOR_PARTE_VARIABILE = &H800080
Public Const FORECOLOR_PARTE_FISSA = vbRed
Public Const FORECOLOR_CAMPI_CALC = vbBlack
Public Const BACKCOLOR_CAMPI_CALC = &HE0E0E0

'Paramtri sistema
Public Const PARAM_EURO = "EUR"
Public Const PARAM_LIRE = "LIT"
Public Const PARAM_GIORNIDEMO = "GIORNIDEMO"


'Il codice livello del personale è fisso LIVn dove n è il livello
'n=2 to 7
'n=8 settimo quadro
'n=9 Ottavo
'n=10 Ottavo Quadro
'n=11 Amministratore
'n=12 Manager
Public Const COD_LIVELLO = "LIV"

'chiave in un treeview , per identificare il livello
'es.KEYLIV1 = livello 1
'   KEYLIV1 & KEYLIV2 = livello 2
Public Const KEYLIV1 = "@KEYLIV1@"
Public Const KEYLIV2 = "@KEYLIV2@"


Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Stringa di manutenzione utilizzata dai servizi di supporto tecnico
End Type


Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Sub Main()
    Dim bAttComune As Boolean

    On Error Resume Next

    frmSplash.Show
    DoEvents
    If Not ApriConnessione Then
        MDIMain.Show
        MDIMain.mnuPathDB_Click
        If gDBC Is Nothing Then GoTo mainERR
    End If
    
    gGiorniScandenzaDemo = GetParamValue("1900", vbNullString, PARAM_GIORNIDEMO)
    If gGiorniScandenzaDemo = -1 Then gGiorniScandenzaDemo = 30
    
    On Error GoTo mainERR
    MDIMain.Show
    Set gJRO = CreateObject("jro.JetEngine")
    DoEvents
    
    Unload frmSplash
    
    'Verifico se è la prima volta che il PGM viene avviato se si chiedo il comune
    bAttComune = ComuneGiaAttivato
    
    'attivo i menu del superUtente
    MDIMain.mnuProtezione.Visible = gComuneAttivo = ALL_COMUNI
    MDIMain.mnuImportazione.Visible = gComuneAttivo = ALL_COMUNI
    
    
    gValutaDefault = GetSetting(App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_VALUTA, "")
    If gValutaDefault = vbNullString Then
        gValutaDefault = VALUTA_EURO
        SaveSetting App.Title, SEZ_ULTIMI_VALORI, KEY_ULTIMI_VALORI_VALUTA, gValutaDefault
    End If
    

    
    If Not bAttComune Then frmLogin.Show vbModal
    
    If gErrore Then
        Err.Raise ID_ERR_PERSONALE, "MAIN", "Attenzione attivazione comune non eseguita"
    Else
        MDIMain.attivaTimer
    End If
    
    
    Exit Sub
mainERR:
    
    Unload frmSplash
    DoEvents
    If Err.Number > 0 Then MsgBox Error$, vbCritical + vbOKOnly, App.Title
    If Not MDIMain Is Nothing Then Unload MDIMain
End Sub

Function ComuneGiaAttivato() As Boolean
    'Questa funzione attiva il programma di bollettazione automatica per un comune
    'che avvia il pgm la prima volta
    '
    'Input:Nessuno
    'Output:True  Comune già attivato
    '       False Comune da attivare
    Dim rs As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ComuneGiaAttivatoErr
    Set rs = gDBC.Execute("Select Attivazione.CodComune as CodComune,Comuni.Comune as comune from Attivazione Left Join Comuni On Attivazione.CodComune=Comuni.CodIstat")
    
    ComuneGiaAttivato = False
    If Not rs.EOF Then
        ComuneGiaAttivato = True
        If rs("CodComune") = ALL_COMUNI Then
            gWhereSuComune = vbNullString
            gComuneAttivo = ALL_COMUNI
        Else
            gWhereSuComune = "Where CodIstat in ("
        End If
        i = 0
        Do While Not rs.EOF
            If rs("CodComune") <> ALL_COMUNI Then
                gWhereSuComune = gWhereSuComune & "'" & rs("CodComune") & "',"
            End If
            gComuneAttivo = rs("CodComune")
            MDIMain.Caption = MDIMain.Caption & " - " & IIf(rs("CodComune") = ALL_COMUNI, rs("CodComune"), rs("Comune"))
            i = i + 1
            rs.MoveNext
        Loop
        If gComuneAttivo <> ALL_COMUNI Then
            gWhereSuComune = Left$(gWhereSuComune, Len(gWhereSuComune) - 1) & ")"
            'se ho più di un comune attivo simulo la ALL_COMUNI per scegliere tra in n attivi durante le fasi del pgm
            If i > 1 Then
                gComuneAttivo = ALL_COMUNI
            End If
        End If
    End If
    Set rs = Nothing
    Exit Function
ComuneGiaAttivatoErr:
    ComuneGiaAttivato = False
    Err.Raise ID_ERR_PERSONALE, "AttivaComune", Error$
End Function


Function AttivaComune(codIstat As String, NomeComune As String) As Boolean
    'Attiva il comune selezionato , inserendo l'indentificativo del comune
    'nella tabella Attivazione
    On Error GoTo AttivaComuneERR
    
    codIstat = Trim$(codIstat)
    AttivaComune = False
    NomeComune = UCase$(NomeComune)
    gDBC.Execute "Insert into Attivazione (CodComune,AnnoRiferimento) values('" & codIstat & "','" & Year(Now) & "')"
    AttivaComune = True
    Exit Function
AttivaComuneERR:
    MsgBox Error$, vbCritical + vbOKOnly, App.Title
End Function


Sub CentraME(frm As Form)
    frm.Move (MDIMain.ScaleWidth - frm.ScaleWidth) / 2, (MDIMain.ScaleHeight - frm.ScaleHeight) / 2
End Sub


Function SetValueInTag(ByVal KEYTAG As String, ByVal Valore As String, sTag As String) As String
    'Cambia il valore a una  Key memorizzata in un tag
    Dim iKey As Integer
    Dim sRight As String
    Dim idSuccKey As Integer
    
    iKey = InStr(1, sTag, KEYTAG)
    If iKey = 0 Then
    'Non esiste la KEY la creo
        sTag = sTag & KEYTAG & "=" & Valore
    Else
        idSuccKey = InStr(iKey + 1, sTag, SEPARATORE)
        If idSuccKey > 0 Then
            sRight = Mid$(sTag, idSuccKey)
        Else
            sRight = vbNullString
        End If
        sTag = Left$(sTag, iKey - 1) & KEYTAG & "=" & Valore & sRight
    End If
    
    SetValueInTag = sTag
End Function



Function GetValueInTag(ByVal KEYTAG As String, sTag As String) As String
    'Prende il valore in una Key memorizzata in Tag
    Dim iKey As Integer
    Dim svalue As String
    Dim iSucc
    
    iKey = InStr(1, sTag, KEYTAG)
    
    If iKey > 0 Then
        iSucc = InStr(iKey + 1, sTag, SEPARATORE)
        
        If iSucc Then
            iSucc = iSucc - iKey - Len(KEYTAG) - 1
        Else
            iSucc = Len(sTag)
        End If
        
        svalue = Mid$(sTag, iKey + Len(KEYTAG) + 1, iSucc)
    Else
        svalue = vbNullString
    End If
    
    GetValueInTag = svalue
End Function

Function GetParamValue(ByVal AnnoRiferimento As String, ByVal CodComune As String, ByVal Parametro As String, Optional ValMax As Variant, Optional ValMin As Variant) As Variant
    Dim rs As Recordset
    On Error GoTo GetParamValueERR
    
    'Legge un il valore di un parametro dalla Tabella Parametri
    Set rs = gDBC.Execute("select * from ParametriSistema where ((AnnoRiferimento='" & AnnoRiferimento & "' and CodIstat='" & CodComune & "') or CodIstat='ALL') and Parametro='" & Parametro & "' order by CodIstat")
    If rs.EOF Then
        GetParamValue = -1
    Else
        GetParamValue = rs("Valore")
        ValMax = rs("ValoreMax")
        ValMin = rs("ValoreMin")
    End If
    
    Exit Function
GetParamValueERR:
    Debug.Print "GetParamValue" & Error$
End Function


Function LetParamValue(ByVal AnnoRiferimento As String, ByVal CodComune As String, ByVal Parametro As String, ByVal LivelloComunale As Boolean, ByVal Valore As String, Optional ValMin As String = -1, Optional ValMax As String = -1) As Boolean
    Dim rs As Recordset
    On Error GoTo LetParamValueERR
    
    'Se LivelloComunale = False il valore del parametro è uguale per tutti i comuni cioè
    'è a livello Italiano
    
    
    If Not LivelloComunale Then
        CodComune = "ALL"
        gDBC.Execute ("Update ParametriSistema set valore='" & Valore & "' where CodIstat='" & CodComune & "' and annoriferimento='1900' and Parametro='" & Parametro & "'")
    Else
        'Legge un il valore di DEFAULT dalla Tabella Parametri la prima volta
        Set rs = gDBC.Execute("select * from ParametriSistema where ((AnnoRiferimento='" & AnnoRiferimento & "' and CodIstat='" & CodComune & "') or CodIstat='ALL') and Parametro='" & Parametro & "' order by CodIstat")
        
        
        If rs(0) = "ALL" Then
            'Setta un valore di un parametro nella Tabella Parametri
            If Valore = vbNullString Then
                Valore = rs("Valore")
            ElseIf Val(Valore) = 0 Then
                Valore = rs("Valore")
            End If
            gDBC.Execute ("insert into ParametriSistema (CodIstat,AnnoRiferimento,Parametro,Descrizione,valore,valoreMin,ValoreMax,Privata) values('" & CodComune & "','" & AnnoRiferimento & "','" & Parametro & "','" & rs("Descrizione") & "','" & Valore & "','" & rs("ValoreMin") & "','" & rs("ValoreMax") & "'," & FormatFLD(rs("Privata"), adBoolean) & ")")
        Else
            gDBC.Execute ("Update ParametriSistema set valore='" & Valore & "' where AnnoRiferimento='" & AnnoRiferimento & "' and CodIstat='" & CodComune & "' and Parametro='" & Parametro & "'")
        End If
    End If
    Set rs = Nothing
    
    Exit Function
LetParamValueERR:
End Function



Function GetValueOfKeyTVW(ByVal KEYNODE As String, sKeyNode As String) As String
    'Prende il valore in una Key
    Dim iKey As Integer
    Dim svalue As String
    Dim iSucc As Integer
    
    iKey = InStr(1, sKeyNode, KEYNODE)
    
    If iKey > 0 Then
        
        iSucc = InStr(iKey + Len(KEYNODE), sKeyNode, Left$(KEYNODE, 1))
    
        If iSucc > 0 Then
            iSucc = iSucc - iKey - Len(KEYNODE)
        Else
            iSucc = Len(sKeyNode)
        End If
        
        svalue = Mid$(sKeyNode, iKey + Len(KEYNODE), iSucc)
    Else
        svalue = vbNullString
    End If
    
    GetValueOfKeyTVW = svalue
End Function

Public Sub Info(txt As String)
    Dim i As Integer
    
    
    If txt = "" Then
        Erase gInfoTxt
        Set MDIMain.StatusBar1.Panels(1).Picture = Nothing
        MDIMain.StatusBar1.Panels(2).Text = ""
        Exit Sub
    End If
    i = -1
    On Error Resume Next
    i = UBound(gInfoTxt)
    ReDim Preserve gInfoTxt(i + 1)
    On Error GoTo 0
    gInfoTxt(i + 1) = txt
    MDIMain.StatusBar1.Panels(2).Text = gInfoTxt(i + 1)
    
    If i + 1 > 0 Then
        Set MDIMain.StatusBar1.Panels(1).Picture = LoadResPicture("INFO", vbResIcon)
        gidxInfo = i + 1
    Else
        Set MDIMain.StatusBar1.Panels(1).Picture = Nothing
    End If
    
    
    
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
    gPathDB = gPathDB & "\StudioAmbiente2002.mdb"
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
                    Err.Raise .Errors(0).NativeError, "OPEN_DB", .Errors(0).description
                End If
                .Errors.Clear
            End If
        Loop While password <> "" And .State <> adStateOpen
    End With
    
    
    gPasswordDB = password
    On Error GoTo ApriConnessioneERR
    'Provider=Microsoft.Jet.OLEDB.4.0;Password="";User ID=Admin;Data Source=C:\sviluppi\GFAmbienteEuro\StudioAmbiente2000.mdb;Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False
    If gDBC.State <> adStateOpen Then
        Err.Raise 9999, "DB OPEN", "Attenzione connessione DB mancante !"
    End If
    
    Set gADOX.ActiveConnection = gDBC
    ApriConnessione = True
    Exit Function
ApriConnessioneERR:
    ChiudiConnessione
End Function

Sub ChiudiConnessione()
    Set gADOX.ActiveConnection = Nothing
    Set gADOX = Nothing
    If Not gDBC Is Nothing Then
        If gDBC.State = adStateOpen Then gDBC.Close
        Set gDBC = Nothing
    End If
End Sub

Function SimboloValuta(valuta As String) As String
    Select Case valuta
        Case VALUTA_EURO
            SimboloValuta = valuta
        Case VALUTA_LIRE
            SimboloValuta = valuta
        Case Else
            SimboloValuta = gValutaDefault
    End Select
End Function


Function PublicMDBonTheWeb(ssql As String, tabella As String, LocalHost As Boolean, sErr As String) As Long
    Dim rs As ADODB.Recordset
    Dim xmlDoc As MSXML2.DOMDocument
    Dim xmlhttp As New MSXML2.xmlhttp
    Dim xmldomTab As New MSXML2.DOMDocument
    Dim sIPAddress As String
    '-1 Errore
    '0 OK
    '1 No dati
    
    
    On Error GoTo PublicMDBonTheWebERR
    sErr = "Errore Generico"
    PublicMDBonTheWeb = -1
    Set rs = gDBC.Execute(ssql)

    If Not rs.EOF Then
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        
        xmlDoc.async = False
        rs.save xmlDoc, adPersistXML
    Else
        If rs.State = adStateOpen Then rs.Close
        sErr = "No Dati"
        PublicMDBonTheWeb = 1
        Exit Function
    End If

    
    xmldomTab.loadXML "<Tabella>" & tabella & "</Tabella>"
    
    xmlDoc.childNodes(0).insertBefore xmldomTab.selectSingleNode("Tabella"), xmlDoc.childNodes(0).childNodes(0)
    
    If LocalHost Then
        sIPAddress = "localhost/tariffaE"
    Else
        sIPAddress = GetParamValue("1900", vbNullString, "HOMEPAGE")
    End If
    
    xmlhttp.Open "POST", "http://" & sIPAddress & "/process.asp", False ', "asgrignoli", "z2yy77"
    
    Dim t As Double
    t = Timer
    xmlhttp.send xmlDoc
    Debug.Print Format$(Timer - t, "###.###")
    
    If InStr(1, xmlhttp.responseText, "Successo") Then
        PublicMDBonTheWeb = 0
        sErr = "Completato"
    Else
        sErr = xmlhttp.responseText & " - " & xmlhttp.statusText
    End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    Set xmlhttp = Nothing
    Set xmlDoc = Nothing
    
    
    
    Exit Function
PublicMDBonTheWebERR:
    sErr = Error$
End Function




Public Function OsVersion() As Long
    Dim os As OSVERSIONINFO
    Dim retval As Long
    os.dwOSVersionInfoSize = Len(os)
    retval = GetVersionEx(os)  ' read Windows's version information
    
'    MsgBox "Windows version number:" & os.dwMajorVersion & Chr(Asc(".")) & os.dwMinorVersion
'    MsgBox "OS Version Info Size = " & os.dwOSVersionInfoSize
'    MsgBox "BuildNumber = " & os.dwBuildNumber
'    MsgBox "Platform ID = " & os.dwPlatformId 'Note If ID =0 win 3.x , ID=1 win9x and ID =2 WINNT
'    MsgBox "CSD Version = " & os.szCSDVersion
    
    OsVersion = Val(os.dwMajorVersion & "." & os.dwMinorVersion)
    
End Function
 
Public Function IsDevEnv() As Boolean
 
    Dim strFileName$
    Dim lngCount&
    
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255&)
    strFileName = Left(strFileName, lngCount)
    
    IsDevEnv = (UCase(Right(strFileName, 7)) Like "VB?.EXE")
 
End Function
 
 
'replace code in my first post with this, the previous code failed to send a key on the first call
Public Sub Sendkeys(Text$, Optional wait As Boolean = False)
 
    Static init As Boolean, IsIDEUnderVista As Boolean, WshShell As Object
    
    'wrapper for Sendkeys which does not cause an Error 70 in the IDE under Windows Vista
    'WshShell does not get set to Nothing before App ends but any problems related to that only happen in the IDE
    'Errors due to WScript disablement on the OS can only happen under the IDE
    'Extra overhead for the WshShell object is only required under the IDE
 
    If Not init Then
        If IsDevEnv() Then
            IsIDEUnderVista = (OsVersion() >= 6)
            If IsIDEUnderVista Then Set WshShell = CreateObject("WScript.Shell")
        End If
        init = True
    End If
    
    If Not IsIDEUnderVista Then
        VBA.Sendkeys Text$, wait
    Else
        WshShell.Sendkeys Text$, wait
    End If
    
End Sub

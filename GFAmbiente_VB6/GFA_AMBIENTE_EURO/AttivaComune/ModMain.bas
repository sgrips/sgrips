Attribute VB_Name = "ModMain"
Option Explicit

Public gDBC As ADODB.Connection
Public gErrore As Boolean
Public gComuneAttivo As String
Public Const ALL_COMUNI = "ALLCOM"
Public Const ID_ERR_PERSONALE = 10000
Public gWhereSuComune As String
Public gWhereFiltroSuComune As String
Public gPathDB As String
Public gPasswordDB As String
Public Const PARAM_GIORNIDEMO = "GIORNIDEMO"

Sub Main()
    
    On Error GoTo mainERR

    'come prima cosa mi copia il dbVuoto
    If Dir$(App.Path & "\StudioAmbiente2002vuoto.mdb", vbArchive) <> "" Then
        FileCopy App.Path & "\StudioAmbiente2002vuoto.mdb", App.Path & "\StudioAmbiente2002.mdb"
    Else
        MsgBox "non trovato " & App.Path & "\StudioAmbiente2000vuoto.mdb", vbCritical + vbOKOnly
        
    End If
    
    Dim bAttComune As Boolean

    On Error Resume Next

    frmSplash.Show
    DoEvents
    
    If Not ApriConnessione Then
        If gDBC Is Nothing Then GoTo mainERR
    End If
    
    
        
    On Error GoTo mainERR
    
    
    
    Unload frmSplash
    
   
    'Verifico se è la prima volta che il PGM viene avviato se si chiedo il comune
    bAttComune = ComuneGiaAttivato
    
    
    If Not bAttComune Then frmLogin.Show vbModal
    If gErrore Then Err.Raise ID_ERR_PERSONALE, "MAIN", "Attenzione attivazione comune non eseguita"

    
    Exit Sub
mainERR:
    
    Unload frmSplash
    DoEvents
    If Err.Number > 0 Then MsgBox Error$, vbCritical + vbOKOnly, App.Title
    

    
End Sub


Function AttivaComune(CodIstat As String, NomeComune As String) As Boolean
    'Attiva il comune selezionato , inserendo l'indentificativo del comune
    'nella tabella Attivazione
    On Error GoTo AttivaComuneERR
    
    CodIstat = Trim$(CodIstat)
    AttivaComune = False
    NomeComune = UCase$(NomeComune)
    gDBC.Execute "Insert into Attivazione (CodComune,AnnoRiferimento) values('" & CodIstat & "','" & Year(Now) & "')"
    AttivaComune = True
    Exit Function
AttivaComuneERR:
    MsgBox Error$, vbCritical + vbOKOnly, App.Title
End Function



Function ComuneGiaAttivato() As Boolean
    'Questa funzione attiva il programma di bollettazione automatica per un comune
    'che avvia il pgm la prima volta
    '
    'Input:Nessuno
    'Output:True  Comune già attivato
    '       False Comune da attivare
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim s As String
    
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
            
            s = s + IIf(rs("CodComune") = ALL_COMUNI, rs("CodComune"), rs("Comune") & vbCrLf)
            
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
        
        MsgBox s, vbExclamation + vbOKOnly
        
    End If
    Set rs = Nothing
    Exit Function
ComuneGiaAttivatoErr:
    ComuneGiaAttivato = False
    Err.Raise ID_ERR_PERSONALE, "AttivaComune", Error$
End Function



Function ApriConnessione(Optional exclusive As Boolean = False) As Boolean
    Dim password As String
    Dim numPassword As Integer
    On Error GoTo ApriConnessioneERR
    
    If Not gDBC Is Nothing Then
        ApriConnessione = True
        Exit Function
    End If
    ApriConnessione = False
    gPathDB = App.Path & "\StudioAmbiente2002.mdb"
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
        Loop While password <> "" And .State <> adStateOpen
    End With
    
    
    gPasswordDB = password
    On Error GoTo ApriConnessioneERR
    'Provider=Microsoft.Jet.OLEDB.4.0;Password="";User ID=Admin;Data Source=C:\sviluppi\GFAmbienteEuro\StudioAmbiente2000.mdb;Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False
    If gDBC.State <> adStateOpen Then
        Err.Raise 9999, "DB OPEN", "Attenzione connessione DB mancante !"
    End If
    

    ApriConnessione = True
    Exit Function
ApriConnessioneERR:
    ChiudiConnessione
End Function


Sub ChiudiConnessione()
    If Not gDBC Is Nothing Then
        If gDBC.State = adStateOpen Then gDBC.Close
        Set gDBC = Nothing
    End If
End Sub


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


Function FormatFLD(ByVal s As String, ByVal Tipo As DataTypeEnum) As String
    'Formatta il campo a seconda del tipo
    Dim i As Integer

    If (s = vbNullString) Then
        FormatFLD = "NULL"
    Else
        Select Case Tipo
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


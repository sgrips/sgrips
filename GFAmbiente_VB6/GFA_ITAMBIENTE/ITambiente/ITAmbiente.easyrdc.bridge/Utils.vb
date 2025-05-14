Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Text


Public Class SettingsRDC
    Const APPLICACTION_SETTING = "Software\VB and VBA Program Settings\IsolaEcologica\impostazioni"
    Const DATABASE_SETTING = "Software\VB and VBA Program Settings\IsolaEcologica\database"

    Dim _settings As Hashtable
    Shared _settingInstace As SettingsRDC = Nothing

    Public Const SETTING_SOCIETA = "SOCIETA"
    Public Const SETTING_IMPIANTO = "IMPIANTO"
    Public Const SETTING_ECOPUNTO_CAUSALE_MOVIMENTO_OUT = "CODCAUSALEMOVECOPUNTOOUT"
    Public Const SETTING_ECOPUNTO_CAUSALE_MOVIMENTO_IN = "CODCAUSALEMOVECOPUNTOIN"
    Public Const SETTING_ECOPUNTO_GENERA_GIACENZA = "FLGMOVECOPUNTOGENMOV"
    Public Const SETTING_PERCORSO_IMPORTAZIONE = "PERCORSOIMP"
    Public Const SETTING_PERCORSO_ESPORTAZIONE = "PERCORSOESP"
    Public Const SETTING_SCONTRINO_ATTIVA_DICHIARAZIONE = "STAMPASCONTRINOFLGDICHIARAZIONE"
    Public Const SETTING_SCONTRINO_DICHIARAZIONE = "STAMPASCONTRINODICHIARAZIONE"
    Public Const SETTING_ULTIMO_UTENTE_CONNESSO = "LASTUSER"
    Public Const SETTING_PERCORSO_APPLICAZIONE = "PERCORSOAPPLICAZIONE"
    Public Const SETTING_LIVELLO_PUNTI_SCUOLE = "PUNTIASSEGNATIPER"


    Public Const SETTING_DATABASE_PATH = "PATH"

    Private Shared instanceOBJ As New Object
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Instance() As SettingsRDC
        SyncLock instanceOBJ

            If _settingInstace Is Nothing Then
                _settingInstace = New SettingsRDC
            End If

            Return _settingInstace

        End SyncLock
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
        Try

            Dim rkey = Registry.CurrentUser.OpenSubKey(APPLICACTION_SETTING)
            Dim settingValues = rkey.GetValueNames()

            _settings = New Hashtable
            For Each a In settingValues
                _settings.Add(a, rkey.GetValue(a))
            Next

            rkey = Registry.CurrentUser.OpenSubKey(DATABASE_SETTING)
            settingValues = rkey.GetValueNames()

            For Each a In settingValues
                _settings.Add(a, rkey.GetValue(a))
            Next
        Catch ex As Exception

        End Try


    End Sub


    ''' <summary>
    '''GGREMINDERGIACGESTIONALE
    '''SOCIETA
    '''IMPIANTO
    '''STAMPAPESATAANTEPRIMA
    '''STAMPAPESATASTAMPANTE
    '''STAMPAPESATANRCOPIE
    '''STAMPAPESATAMARGSX
    '''STAMPAPESATAMARGDX
    '''STAMPAPESATAMARGUP
    '''STAMPAPESATAMARGDWN
    '''STAMPAREGISTROANTEPRIMA
    '''STAMPAREGISTROSTAMPANTE
    '''STAMPAREGISTRONRCOPIE
    '''STAMPAREGISTROMARGSX
    '''STAMPAREGISTROMARGDX
    '''STAMPAREGISTROMARGUP
    '''STAMPAREGISTROMARGDWN
    '''STAMPAFORMULARIOANTEPRIMA
    '''STAMPAFORMULARIOSTAMPANTE
    '''STAMPAFORMULARIONRCOPIE
    '''STAMPAFORMULARIOMARGSX
    '''STAMPAFORMULARIOMARGDX
    '''STAMPAFORMULARIOMARGUP
    '''STAMPAFORMULARIOMARGDWN
    '''CODCAUSALEMOVDAUSCITAPIU
    '''CODUTENTEMOVDAUSCITAPIU
    '''CODIMMOBILEMOVDAUSCITAPIU
    '''CODUSERMOVDAUSCITAPIU
    '''CODCAUSALEMOVECOPUNTOIN
    '''CODCAUSALEMOVECOPUNTOOUT
    '''FLGMOVECOPUNTOGENMOV
    '''FLGCODPESAECOPUNTO
    '''FLGECOPUNTODATAMOVIMENTOINS
    '''BONUSFILE
    '''BONUSFILETYPE
    '''BONUSFILESEPARATORE
    '''LASTUSER
    '''FATTUREFILE
    '''FATTUREFILETYPE
    '''FATTUREFILESEPARATORE
    '''MOVIMENTIFILE
    '''MOVIMENTIFILETYPE
    '''MOVIMENTIFILESEPARATORE
    '''CODICICERFILE
    '''CODICICERFILETYPE
    '''CODICICERFILESEPARATORE
    '''UTENTIFILE
    '''UTENTIFILETYPE
    '''UTENTIFILESEPARATORE
    '''PERCORSOIMP
    '''PERCORSOESP
    '''GESTIONEIMMOBILI
    '''STAMPASCONTRINOANTEPRIMA
    '''STAMPASCONTRINOSTAMPANTE
    '''STAMPASCONTRINONRCOPIE
    '''STAMPASCONTRINOMARGSX
    '''STAMPASCONTRINOMARGDX
    '''STAMPASCONTRINOMARGUP
    '''STAMPASCONTRINOMARGDWN
    '''NOTIFICHE
    '''STAMPASCONTRINOFLGDICHIARAZIONE
    '''STAMPASCONTRINODICHIARAZIONE
    ''' </summary>
    ''' <param name="propertyName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPropertyValue(ByVal propertyName As String) As Object


        If _settings IsNot Nothing AndAlso _settings.ContainsKey(propertyName) Then
            Return _settings.Item(propertyName)
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Ritorna la stringa di connessione al database 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetConnectionString() As String
        Dim pathDB = SettingsRDC.Instance.GetPropertyValue(SettingsRDC.SETTING_DATABASE_PATH)
        Dim pwdDB = Crittografy.GetPasswordDB(0, SettingsRDC.Instance.GetPropertyValue(SettingsRDC.SETTING_PERCORSO_APPLICAZIONE))

        Return String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Database Password={1}", pathDB, pwdDB)
    End Function

End Class


Public Class Crittografy
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function GetPrivateProfileString(ByVal lpAppName As String, ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As StringBuilder, _
        ByVal nSize As Int32, _
        ByVal lpFileName As String) As Int32
    End Function



    'Set to True to make the password case-sensitive
#Const CASE_SENSITIVE_PASSWORD = False
    Const password = "gfambiente"

    ''' <summary>
    ''' Decriptare un testo passato
    ''' </summary>
    ''' <param name="strText">Testo da decriptare</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTextDecrypt(ByVal strText As String) As String
        strText = DecryptText(strText, password)
        GetTextDecrypt = DecryptText(strText, password)

    End Function

    ''' <summary>
    ''' Ritorna un testo criptato
    ''' </summary>
    ''' <param name="strText">Testo da criptare</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetTextEncrypt(ByVal strText As String) As String
        strText = EncryptText(strText, password)
        GetTextEncrypt = EncryptText(strText, password)

    End Function



    'Encrypt text
    Private Shared Function DecryptText(ByVal strText As String, ByVal strPwd As String)
        Dim i As Integer, c As Integer
        Dim strBuff As String = String.Empty

#If Not CASE_SENSITIVE_PASSWORD Then

        'Convert password to upper case
        'if not case-sensitive
        strPwd = UCase$(strPwd)

#End If

        'Encrypt string
        If Len(strPwd) Then
            For i = 1 To Len(strText)
                c = Asc(Mid$(strText, i, 1))
                c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
                strBuff = strBuff & Chr(c And &HFF)
            Next i
        Else
            strBuff = strText
        End If
        Return strBuff
    End Function


    Private Shared Function EncryptText(ByVal strText As String, ByVal strPwd As String)
        Dim i As Integer, c As Integer
        Dim strBuff As String = String.Empty

#If Not CASE_SENSITIVE_PASSWORD Then

        'Convert password to upper case
        'if not case-sensitive
        strPwd = UCase$(strPwd)

#End If

        'Decrypt string
        If Len(strPwd) Then
            For i = 1 To Len(strText)
                c = Asc(Mid$(strText, i, 1))
                c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
                strBuff = strBuff & Chr(c And &HFF)
            Next i
        Else
            strBuff = strText
        End If
        Return strBuff
    End Function

    ''' <summary>
    ''' Ritorna la password del data base decriptata
    ''' </summary>
    ''' <param name="numPass">ID password da restituire</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetPasswordDB(ByVal numPass As Integer, ByVal applicationPath As String) As String
        Dim retString As New StringBuilder(Space(1024))
        Dim ret As String
        Dim l As Long


        l = GetPrivateProfileString("PASSWORD", "password" & CStr(numPass), "", retString, retString.Length, applicationPath & "\config.ini")
        If (l > 0) Then
            ret = retString.ToString()
        Else
            ret = "" : Return Nothing
        End If
        numPass = numPass + 1


        Return GetTextDecrypt(Trim$(ret))
    End Function





End Class

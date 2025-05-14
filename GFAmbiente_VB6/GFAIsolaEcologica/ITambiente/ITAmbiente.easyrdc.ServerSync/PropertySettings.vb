Imports System.ComponentModel
Imports System.Windows.Forms.Design
Imports System.Drawing.Design
Imports ITAmbiente.easyrdc.Sync.FolderBrowserX
#Region "FolderName Browser"
''' <summary>
''' Classe per la gestione del Browsing Dir
''' </summary>
''' <remarks></remarks>
Friend Class FolderBrowserX
    Inherits FolderNameEditor

    Private _Description As String

    Protected Overrides Sub InitializeDialog(ByVal folderBrowser As System.Windows.Forms.Design.FolderNameEditor.FolderBrowser)
        folderBrowser.Description = _Description

        MyBase.InitializeDialog(folderBrowser)
    End Sub


    Public Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
        Dim attDescription As FolderDescriptionAttribute = context.PropertyDescriptor.Attributes(GetType(FolderDescriptionAttribute))

        If attDescription IsNot Nothing Then _Description = attDescription.Description

        Return MyBase.GetEditStyle(context)
    End Function

    Public Class FolderDescriptionAttribute : Inherits Attribute
        Private _description As String
        Property Description() As String
            Get
                Return _description
            End Get
            Set(ByVal value As String)
                _description = value
            End Set
        End Property

        Public Sub New(ByVal _desription As String)
            Me._description = _desription
        End Sub

    End Class
End Class

#End Region

Public Enum ConnectionModeEnum
    PIPE = 0
    SHAREDFOLDER = 1
End Enum


Public Class PropertySettings
    Dim _forderNameSync As String
    Dim _timeStartSync As TimeSpan
    Dim _timeEndSync As TimeSpan
    Dim _pathDataBase As String
    Dim _fileNamePattern As String
    Dim _passwordDB As String
    Dim _modeConnect As ConnectionModeEnum
    Dim _sharedFolderName As String

    Public Sub New()
        _forderNameSync = My.Settings.FolderNameSync
        _timeStartSync = My.Settings.TimeStartSync
        _timeEndSync = My.Settings.TimeEndSync
        _pathDataBase = My.Settings.PathDataBase
        _passwordDB = My.Settings.PasswordDB
        _modeConnect = My.Settings.ConnectionMode
        _fileNamePattern = My.Settings.FileNamePattern
        _sharedFolderName = My.Settings.SharedFolderName
    End Sub


    Public Sub Save()
        My.Settings.FolderNameSync = _forderNameSync
        My.Settings.TimeStartSync = _timeStartSync
        My.Settings.TimeEndSync = _timeEndSync
        My.Settings.PathDataBase = _pathDataBase
        My.Settings.PasswordDB = _passwordDB
        My.Settings.FileNamePattern = _fileNamePattern
        My.Settings.ConnectionMode = _modeConnect
        My.Settings.SharedFolderName = _sharedFolderName
        My.Settings.Save()
    End Sub

  


    <CategoryAttribute("Sync"), DisplayName("Ora Fine")> _
    Property TimeEndSync() As TimeSpan
        Get
            Return _timeEndSync
        End Get
        Set(ByVal value As TimeSpan)
            _timeEndSync = value
        End Set
    End Property

    <CategoryAttribute("Sync"), DisplayName("Ora Inizio")> _
    Property TimeStartSync() As TimeSpan
        Get
            Return _timeStartSync
        End Get
        Set(ByVal value As TimeSpan)
            _timeStartSync = value
        End Set
    End Property


    <Description("Percorso dove avviene la sincronizzazione dei dati tra i dispositivi mobile ed il pc master"), _
           DefaultValue("HOME"), _
           CategoryAttribute("Sync"), _
           DisplayName("Directory Sync"), _
           FolderDescriptionAttribute("Indicare il percorso da monitorare in cui vengono inseriti i file 'Movimenti_XX.csv' e totali per badge 'Movimenti_XXTot.xml'"), _
           Editor(GetType(FolderBrowserX), GetType(UITypeEditor)), _
           Browsable(True)> _
    Property FolderNameSync() As String
        Get
            Return _forderNameSync
        End Get
        Set(ByVal value As String)
            _forderNameSync = value
        End Set
    End Property

    <Description("Percorso del file DataBase Master"), _
           DefaultValue("HOME"), _
           CategoryAttribute("Data Base"), _
           DisplayName("File MDB"), _
           FolderDescriptionAttribute("Indicare il percorso dove si trova il file MDB del database Master."), _
           Editor(GetType(FolderBrowserX), GetType(UITypeEditor)), _
           Browsable(True)> _
       Property PathDataBase() As String
        Get
            Return _pathDataBase
        End Get
        Set(ByVal value As String)
            _pathDataBase = value
        End Set
    End Property

    <Description("Indicare la password per il database Master"), _
    PasswordPropertyText(True), _
    DisplayName("Password"), _
    CategoryAttribute("Data Base")> _
    Property PasswordDB() As String
        Get
            Dim cPWD = Util.Security.Cryptography.Decrypt(_passwordDB, "ITAmbiente")
            Return Strings.Space(10 - cPWD.Length)
        End Get
        Set(ByVal value As String)
            _passwordDB = Util.Security.Cryptography.Encrypt(Trim$(value), "ITAmbiente")
        End Set
    End Property

    <CategoryAttribute("Sync"), DisplayName("File Pattern")> _
  Property FileNamePattern() As String
        Get
            Return _fileNamePattern
        End Get
        Set(ByVal value As String)
            _fileNamePattern = value
        End Set
    End Property


    <DisplayName("Modo Connessione"), _
    CategoryAttribute("Server")> _
    Property ConnectionMode() As ConnectionModeEnum
        Get
            Return _modeConnect
        End Get
        Set(ByVal value As ConnectionModeEnum)
            _modeConnect = value
        End Set
    End Property

    <DisplayName("Client Cartella Condivisa"), _
        CategoryAttribute("Server")> _
        Property SharedFolderName() As String
        Get
            Return _sharedFolderName
        End Get
        Set(ByVal value As String)
            _sharedFolderName = value
        End Set
    End Property



End Class
